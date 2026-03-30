# -*- coding: utf-8 -*-
"""
Mover Caminos de Circuito – PyRevit
====================================
Replica el Dynamo "mover caminos de circuito v3.dyn".

Correcciones aplicadas vs. versión anterior:
  - Punto 0 del camino = CONECTOR del cuadro (panel), no el equipo.
  - El camino resultante se ortogonaliza (sin diagonales XY+Z simultáneas).
  - Se eliminan puntos consecutivos demasiado cercanos.
  - GetCircuitPath() se convierte correctamente a lista Python.
"""

import clr
import math
import heapq
from collections import defaultdict

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    LocationCurve, LocationPoint,
    Transaction, XYZ,
)

from pyrevit import revit, script as pvscript

doc    = revit.doc
logger = pvscript.get_logger()

# ─────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────
MM = 1.0 / 304.8              # mm → pies (unidad interna Revit)

DIST_INTERMEDIA  = 200.0 * MM  # Paso de discretización en bandejas
DIST_SALTO       = 300.0 * MM  # Salto máximo entre bandejas
TOLERANCIA       = 8000.0 * MM # Radio para conectar inicio/fin al grafo
MIN_DIST_DEFAULT = 1.0 * MM    # Respaldo si no se puede leer la tolerancia real de Revit
AXIS_TOL         = 1e-4

PARAM_CIRC_NUM  = u"Número de circuito"
PARAM_UBICACION = u"ubicación"
PARAM_COMENT    = u"Comentarios"

DEBUG = False


# ═══════════════════════════════════════════════════════════
# PASO 1 – Bandejas portacables
# ═══════════════════════════════════════════════════════════

def get_tray_lines(doc):
    lines = []
    tol = doc.Application.ShortCurveTolerance

    for t in (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_CableTray)
              .WhereElementIsNotElementType()):
        loc = t.Location
        if isinstance(loc, LocationCurve):
            c = loc.Curve
            if c.Length > tol:
                lines.append((c.GetEndPoint(0), c.GetEndPoint(1)))

    for f in (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_CableTrayFitting)
              .WhereElementIsNotElementType()):
        try:
            conns = list(f.MEPModel.ConnectorManager.Connectors)
            for i in range(len(conns)):
                for j in range(i + 1, len(conns)):
                    p0, p1 = conns[i].Origin, conns[j].Origin
                    if p0.DistanceTo(p1) > tol:
                        lines.append((p0, p1))
        except Exception:
            pass
    return lines


# ═══════════════════════════════════════════════════════════
# PASO 2 – Grafo de Dijkstra
# ═══════════════════════════════════════════════════════════

def lerp(a, b, t):
    return XYZ(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t, a.Z + (b.Z - a.Z) * t)

def xkey(pt, prec=4):
    return (round(pt.X, prec), round(pt.Y, prec), round(pt.Z, prec))

def discretize(p0, p1, step):
    L = p0.DistanceTo(p1)
    if L < 1e-9:
        return [p0]
    n = max(1, int(math.floor(L / step)))
    return [lerp(p0, p1, i / float(n)) for i in range(n + 1)]

def build_graph(lines, step=DIST_INTERMEDIA, jump=DIST_SALTO):
    id2xyz = {}
    seg_keys_list = []

    for p0, p1 in lines:
        pts = discretize(p0, p1, step)
        seg_keys = []
        for pt in pts:
            k = xkey(pt)
            id2xyz.setdefault(k, pt)
            seg_keys.append(k)
        seg_keys_list.append(seg_keys)

    graph = defaultdict(list)

    for seg_keys in seg_keys_list:
        for i in range(len(seg_keys) - 1):
            ka, kb = seg_keys[i], seg_keys[i + 1]
            if ka == kb:
                continue
            d = id2xyz[ka].DistanceTo(id2xyz[kb])
            graph[ka].append((d, kb))
            graph[kb].append((d, ka))

    all_keys = list(id2xyz.keys())
    for i in range(len(all_keys)):
        ka = all_keys[i]
        for j in range(i + 1, len(all_keys)):
            kb = all_keys[j]
            d = id2xyz[ka].DistanceTo(id2xyz[kb])
            if 0 < d <= jump:
                graph[ka].append((d, kb))
                graph[kb].append((d, ka))

    return graph, id2xyz


def connect_point(graph, id2xyz, tray_keys, pt, tol):
    """Conecta pt al nodo más cercano del grafo. Devuelve (key, distancia)."""
    best_d, best_k = float('inf'), None
    for k in tray_keys:
        d = id2xyz[k].DistanceTo(pt)
        if d < best_d:
            best_d, best_k = d, k
    if best_k is None or best_d > tol:
        return None, best_d
    new_k = ("_ext_", round(pt.X, 6), round(pt.Y, 6), round(pt.Z, 6))
    id2xyz[new_k] = pt
    graph[new_k].append((best_d, best_k))
    graph[best_k].append((best_d, new_k))
    return new_k, best_d


def dijkstra(graph, start, end):
    heap = [(0.0, start)]
    dist = defaultdict(lambda: float('inf'))
    dist[start] = 0.0
    prev = {start: None}
    seen = set()
    while heap:
        cost, node = heapq.heappop(heap)
        if node in seen:
            continue
        seen.add(node)
        if node == end:
            break
        for w, nb in graph.get(node, []):
            nc = cost + w
            if nc < dist[nb]:
                dist[nb] = nc
                prev[nb] = node
                heapq.heappush(heap, (nc, nb))
    if dist[end] == float('inf'):
        return None, float('inf')
    path, cur = [], end
    while cur is not None:
        path.append(cur)
        cur = prev.get(cur)
    path.reverse()
    return path, dist[end]


# ═══════════════════════════════════════════════════════════
# ORTOGONALIZACIÓN Y LIMPIEZA DEL CAMINO
# ═══════════════════════════════════════════════════════════

def orthogonalize(pts, tol=AXIS_TOL):
    """
    Garantiza que cada segmento sea horizontal (XY) o vertical (Z).
    Si un segmento es diagonal, inserta un punto intermedio:
    primero mueve en XY, luego en Z.
    """
    if not pts:
        return pts
    result = [pts[0]]
    for curr in pts[1:]:
        prev = result[-1]
        dx = curr.X - prev.X
        dy = curr.Y - prev.Y
        dz = curr.Z - prev.Z
        move_xy = abs(dx) > tol or abs(dy) > tol
        move_z  = abs(dz) > tol
        if move_xy and move_z:
            # Mueve primero en XY (misma Z que el punto anterior), luego sube/baja
            mid = XYZ(curr.X, curr.Y, prev.Z)
            result.append(mid)
        result.append(curr)
    return result


def points_equal(a, b, tol=AXIS_TOL):
    return a.DistanceTo(b) <= tol


def dedupe_consecutive_pts(pts, tol=AXIS_TOL):
    """Elimina duplicados consecutivos sin romper la geometría ortogonal."""
    if not pts:
        return pts
    result = [pts[0]]
    for pt in pts[1:]:
        if not points_equal(pt, result[-1], tol):
            result.append(pt)
    return result


def is_axis_aligned(a, b, tol=AXIS_TOL):
    same_xy = abs(a.X - b.X) <= tol and abs(a.Y - b.Y) <= tol
    same_z  = abs(a.Z - b.Z) <= tol
    return same_xy or same_z


def validate_path_nodes(pts, min_d, tol=AXIS_TOL):
    if len(pts) < 2:
        return False, "La ruta final tiene menos de 2 puntos"

    for i in range(len(pts) - 1):
        a = pts[i]
        b = pts[i + 1]
        seg_len = a.DistanceTo(b)
        if seg_len < min_d:
            return False, "Segmento {}-{} demasiado corto ({:.1f} mm)".format(
                i, i + 1, seg_len / MM)
        if not is_axis_aligned(a, b, tol):
            return False, "Segmento {}-{} no es horizontal/vertical".format(i, i + 1)

    return True, "OK"


def connection_candidates(a, b, tol=AXIS_TOL):
    """Devuelve posibles conexiones ortogonales entre dos puntos."""
    if points_equal(a, b, tol):
        return []

    if is_axis_aligned(a, b, tol):
        return [[a, b]]

    mids = [
        XYZ(b.X, b.Y, a.Z),  # XY y luego Z
        XYZ(a.X, a.Y, b.Z),  # Z y luego XY
    ]

    chains = []
    seen = set()
    for mid in mids:
        chain = dedupe_consecutive_pts([a, mid, b], tol)
        if len(chain) < 2:
            continue
        key = tuple(xkey(pt, 6) for pt in chain)
        if key in seen:
            continue
        seen.add(key)
        chains.append(chain)
    return chains


def choose_connection(a, b, min_d, tol=AXIS_TOL):
    """Elige una conexión ortogonal válida entre dos puntos."""
    candidates = []
    for chain in connection_candidates(a, b, tol):
        ok, _ = validate_path_nodes(chain, min_d, tol)
        if not ok:
            continue
        seg_lengths = [chain[i].DistanceTo(chain[i + 1]) for i in range(len(chain) - 1)]
        candidates.append((len(chain), -min(seg_lengths), chain))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def build_final_path(pt_panel, tray_pts, pt_equipo, min_d, tol=AXIS_TOL):
    """
    Construye una ruta válida para SetCircuitPath.
    Evita que la limpieza elimine nodos pequeños necesarios cerca del panel/equipo.
    """
    tray_pts = dedupe_consecutive_pts(tray_pts, tol)

    if not tray_pts:
        direct = choose_connection(pt_panel, pt_equipo, min_d, tol)
        if direct is None:
            return None, "Sin puntos intermedios y la conexión directa no es válida"
        return direct, "Direct"

    start_options = []
    end_options = []

    for idx, tray_pt in enumerate(tray_pts):
        start_chain = choose_connection(pt_panel, tray_pt, min_d, tol)
        if start_chain is not None:
            start_options.append((idx, start_chain))

        end_chain = choose_connection(tray_pt, pt_equipo, min_d, tol)
        if end_chain is not None:
            end_options.append((idx, end_chain))

    if not start_options:
        return None, "No se encontró conexión válida desde el panel al grafo"
    if not end_options:
        return None, "No se encontró conexión válida desde el grafo al equipo"

    end_map = {idx: chain for idx, chain in end_options}

    for start_idx, start_chain in start_options:
        for end_idx in range(len(tray_pts) - 1, start_idx - 1, -1):
            end_chain = end_map.get(end_idx)
            if end_chain is None:
                continue

            candidate = list(start_chain)
            candidate.extend(tray_pts[start_idx + 1:end_idx + 1])
            candidate.extend(end_chain[1:])

            candidate = dedupe_consecutive_pts(candidate, tol)
            candidate = orthogonalize(candidate, tol)
            candidate = dedupe_consecutive_pts(candidate, tol)

            ok, _ = validate_path_nodes(candidate, min_d, tol)
            if ok:
                return candidate, "TrayPath"

    direct = choose_connection(pt_panel, pt_equipo, min_d, tol)
    if direct is not None:
        return direct, "DirectFallback"

    return None, "No se pudo construir una ruta ortogonal válida"


# ═══════════════════════════════════════════════════════════
# PASO 3 – Endpoints de circuito
#   IMPORTANTE: pt_panel (cuadro) va PRIMERO en SetCircuitPath.
# ═══════════════════════════════════════════════════════════

def elem_location(elem):
    if elem is None:
        return None
    loc = elem.Location
    if isinstance(loc, LocationPoint):
        return loc.Point
    if isinstance(loc, LocationCurve):
        return loc.Curve.Evaluate(0.5, True)
    return None


def get_endpoints(circuit):
    """
    Devuelve (pt_panel, pt_equipo, metodo).
    pt_panel = posición del conector del cuadro (primer punto requerido por SetCircuitPath).
    pt_equipo = posición del equipo conectado (último punto).
    """
    # --- Estrategia 1: GetCircuitPath ---
    try:
        raw = circuit.GetCircuitPath()
        pts = list(raw) if raw else []
        if len(pts) >= 2:
            return pts[0], pts[-1], "GetCircuitPath"
        if len(pts) == 1:
            return pts[0], pts[0], "GetCircuitPath1pt"
    except Exception as e:
        logger.debug("GetCircuitPath: {}".format(e))

    # --- Estrategia 2: Conectores del cuadro ---
    try:
        base_eq = circuit.BaseEquipment
        if base_eq is not None:
            # Buscar el conector del panel que referencia a elementos de este circuito
            circuit_elem_ids = set(e.Id.IntegerValue for e in circuit.Elements)
            cm = base_eq.MEPModel.ConnectorManager
            panel_connector_pt = None
            for conn in cm.Connectors:
                try:
                    for ref in conn.AllRefs:
                        if ref.Owner.Id.IntegerValue in circuit_elem_ids:
                            panel_connector_pt = conn.Origin
                            break
                except Exception:
                    pass
                if panel_connector_pt is not None:
                    break

            # Equipo conectado (primer elemento que no sea el cuadro)
            pt_equipo = None
            for elem in circuit.Elements:
                if elem.Id == base_eq.Id:
                    continue
                pt = elem_location(elem)
                if pt is not None:
                    pt_equipo = pt
                    break

            pt_panel = panel_connector_pt or elem_location(base_eq)

            if pt_panel is not None and pt_equipo is not None:
                return pt_panel, pt_equipo, "Connectors"
    except Exception as e:
        logger.debug("Connector strategy: {}".format(e))

    # --- Estrategia 3: Ubicaciones puras ---
    try:
        base_eq = circuit.BaseEquipment
        pt_panel = elem_location(base_eq)
        pt_equipo = None
        for elem in circuit.Elements:
            if base_eq is not None and elem.Id == base_eq.Id:
                continue
            pt = elem_location(elem)
            if pt is not None:
                pt_equipo = pt
                break
        if pt_panel is not None and pt_equipo is not None:
            return pt_panel, pt_equipo, "Location"
    except Exception as e:
        logger.debug("Location strategy: {}".format(e))

    return None, None, "NoPath"


# ═══════════════════════════════════════════════════════════
# PASO 6 – Parámetros
# ═══════════════════════════════════════════════════════════

def get_param(elem, name):
    p = elem.LookupParameter(name)
    if p is None:
        return None
    try:
        return p.AsString() or p.AsValueString() or ""
    except Exception:
        return ""


def set_param(elem, name, value):
    p = elem.LookupParameter(name)
    if p is None or p.IsReadOnly:
        return False
    try:
        p.Set(value)
        return True
    except Exception as e:
        logger.warning("set_param '{}': {}".format(name, e))
        return False


def build_mech_index(doc):
    idx = {}
    for eq in (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
               .WhereElementIsNotElementType()):
        val = get_param(eq, PARAM_CIRC_NUM)
        if val:
            for part in val.replace(";", ",").split(","):
                k = part.strip()
                if k:
                    idx[k] = eq
    return idx


# ═══════════════════════════════════════════════════════════
# PROGRAMA PRINCIPAL
# ═══════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("MOVER CAMINOS DE CIRCUITO")
    print("=" * 60)

    print("\n[1] Bandejas portacables...")
    lines = get_tray_lines(doc)
    print("    {} segmentos.".format(len(lines)))
    if not lines:
        print("ERROR: sin bandejas.")
        return

    print("\n[2] Construyendo grafo (esto puede tardar)...")
    graph, id2xyz = build_graph(lines)
    tray_keys = list(id2xyz.keys())
    print("    {} nodos.".format(len(tray_keys)))

    print("\n[3] Circuitos eléctricos...")
    circuits = list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_ElectricalCircuit)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    print("    {} circuitos.".format(len(circuits)))
    if not circuits:
        print("Sin circuitos.")
        return

    print("\n[6 prep] Índice de equipos mecánicos...")
    mech_idx = build_mech_index(doc)
    print("    {} entradas.".format(len(mech_idx)))

    min_seg_len = max(getattr(doc.Application, "ShortCurveTolerance", 0.0), MIN_DIST_DEFAULT)

    print("\n[4+5] Calculando y aplicando caminos...")
    print("    TOLERANCIA={:.0f}mm\n".format(TOLERANCIA / MM))

    ok = skip = err = 0
    skipped_log = []   # [(cnum, motivo)]
    error_log   = []   # [(cnum, mensaje)]

    t = Transaction(doc, "Mover Caminos de Circuito")
    t.Start()
    try:
        for circuit in circuits:
            cnum = get_param(circuit, PARAM_CIRC_NUM) or str(circuit.Id.IntegerValue)

            pt_panel, pt_equipo, method = get_endpoints(circuit)

            if pt_panel is None or pt_equipo is None:
                motivo = "Sin endpoints detectados (método={})".format(method)
                if DEBUG: print("  SKIP [{}] {}".format(cnum, motivo))
                skip += 1
                skipped_log.append((cnum, motivo))
                continue

            dist_total = pt_panel.DistanceTo(pt_equipo)
            if dist_total < 1e-6:
                motivo = "Panel y equipo en la misma posición"
                if DEBUG: print("  SKIP [{}] {}".format(cnum, motivo))
                skip += 1
                skipped_log.append((cnum, motivo))
                continue

            # Clonar grafo local para este circuito
            local_graph = defaultdict(list, {k: list(v) for k, v in graph.items()})
            local_xyz   = dict(id2xyz)

            k_panel,  d_panel  = connect_point(local_graph, local_xyz, tray_keys, pt_panel,  TOLERANCIA)
            k_equipo, d_equipo = connect_point(local_graph, local_xyz, tray_keys, pt_equipo, TOLERANCIA)

            if DEBUG:
                print("  [{}] método={} | d_panel→grafo={:.0f}mm | d_equipo→grafo={:.0f}mm".format(
                    cnum, method, d_panel / MM, d_equipo / MM))


            if k_panel is None:
                motivo = "Cuadro a {:.0f}mm del grafo (tol={:.0f}mm)".format(d_panel / MM, TOLERANCIA / MM)
                if DEBUG: print("  SKIP [{}] {}".format(cnum, motivo))
                skip += 1
                skipped_log.append((cnum, motivo))
                continue
            if k_equipo is None:
                eq_mec = mech_idx.get(cnum.strip())
                ubicacion = get_param(eq_mec, PARAM_UBICACION) if eq_mec is not None else "desconocido"
                motivo = "Equipo {} a {:.2f}m".format(
                    ubicacion or "desconocido",
                    d_equipo / MM / 1000.0)
                if DEBUG: print("  SKIP [{}] {}".format(cnum, motivo))
                skip += 1
                skipped_log.append((cnum, motivo))
                continue

            path_keys, cost = dijkstra(local_graph, k_panel, k_equipo)
            if path_keys is None:
                motivo = "Sin camino posible en el grafo de bandejas"
                if DEBUG: print("  SKIP [{}] {}".format(cnum, motivo))
                skip += 1
                skipped_log.append((cnum, motivo))
                continue

            # Construir lista de XYZ
            raw_path = [local_xyz[k] for k in path_keys]
            raw_path[0]  = pt_panel   # Conector del cuadro (PRIMER punto, obligatorio)
            raw_path[-1] = pt_equipo  # Equipo al final

            clean_path, path_mode = build_final_path(
                pt_panel,
                raw_path[1:-1],
                pt_equipo,
                min_seg_len,
                AXIS_TOL
            )

            if clean_path is None:
                motivo = "Ruta inválida tras saneado ({})".format(path_mode)
                if DEBUG: print("  SKIP [{}] {}".format(cnum, motivo))
                skip += 1
                skipped_log.append((cnum, motivo))
                continue

            ok_path, path_msg = validate_path_nodes(clean_path, min_seg_len, AXIS_TOL)
            if not ok_path:
                motivo = "Ruta inválida antes de aplicar: {}".format(path_msg)
                if DEBUG: print("  SKIP [{}] {}".format(cnum, motivo))
                skip += 1
                skipped_log.append((cnum, motivo))
                continue

            try:
                circuit.SetCircuitPath(clean_path)
                ok += 1
                if DEBUG: print("  OK   [{}] {} pts | dist~{:.0f}mm".format(
                    cnum, len(clean_path), cost / MM))
            except Exception as e:
                msg = str(e).splitlines()[0]
                if DEBUG: print("  ERR  [{}] SetCircuitPath: {}".format(cnum, msg))
                err += 1
                error_log.append((cnum, msg))
                continue

            # PASO 6 – Comentarios ← ubicación del equipo mecánico
            equipo_mec = mech_idx.get(cnum.strip())
            if equipo_mec is not None:
                ub = get_param(equipo_mec, PARAM_UBICACION)
                if ub is not None:
                    set_param(circuit, PARAM_COMENT, ub)

        t.Commit()

    except Exception as e:
        try:
            t.RollbackIfPossible()
        except Exception:
            pass
        print("\nERROR FATAL – transacción revertida: {}".format(e))
        import traceback
        traceback.print_exc()

    print("\n" + "=" * 60)
    print("  Actualizados : {}".format(ok))
    print("  Omitidos     : {}".format(skip))
    print("  Errores      : {}".format(err))
    print("=" * 60)

    if skipped_log:
        print("\nCIRCUITOS OMITIDOS:")
        print("  " + "-" * 56)
        for cnum, motivo in skipped_log:
            print("  {:<20} {}".format(cnum, motivo))

    if error_log:
        print("\nCIRCUITOS CON ERROR:")
        print("  {:<20} {}".format("Circuito", "Error"))
        print("  " + "-" * 56)
        for cnum, msg in error_log:
            print("  {:<20} {}".format(cnum, msg))


if __name__ == "__main__":
    main()
