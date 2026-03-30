# -*- coding: utf-8 -*-

__title__ = "Tag Segments A/L"
__author__ = "Juan Achenbach"
__version__ = "Version: 1.0"
__doc__ = """Tag each segment between tees at segment half-length midpoint."""

import sys
import unicodedata

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import (
    HorizontalAlignment,
    ResizeMode,
    SizeToContent,
    Thickness,
    Window,
    WindowStartupLocation,
)
from System.Windows.Controls import Button as WpfButton
from System.Windows.Controls import Label as WpfLabel
from System.Windows.Controls import ListBox as WpfListBox
from System.Windows.Controls import ListBoxItem as WpfListBoxItem
from System.Windows.Controls import Orientation as WpfOrientation
from System.Windows.Controls import StackPanel as WpfStackPanel

from pyrevit import revit
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    LocationCurve,
    PartType,
    Reference,
    TagMode,
    TagOrientation,
    Transaction,
)
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.UI import TaskDialog

doc = revit.doc
view = doc.ActiveView

FAMILY_BASE_NAME = u"CST_TAG Diametro Tubería v{}"
VERSION_MIN = 28
VERSION_MAX = 40
TAG_TYPE_NAME = "1.5 DT"

# Offsets from segment midpoint (meters), by tag type and segment orientation in view.
# Horizontal segment -> use *_H_* ; Vertical segment -> use *_V_*
# Grouped by orientation and sign so A/L pairs are adjacent.
# Horizontal PLUS
A_H_DX_PLUS_M = 0.5
A_H_DY_PLUS_M = 0.5
L_H_DX_PLUS_M = 1.6
L_H_DY_PLUS_M = 1.12

# Horizontal MINUS
A_H_DX_MINUS_M = 0.5
A_H_DY_MINUS_M = -0.5
L_H_DX_MINUS_M = 1.6
L_H_DY_MINUS_M = 0.875

# Vertical PLUS
A_V_DX_PLUS_M = 1
A_V_DY_PLUS_M = 0.5
L_V_DX_PLUS_M = 1.9
L_V_DY_PLUS_M = 0.65

# Vertical MINUS
A_V_DX_MINUS_M = 0.5
A_V_DY_MINUS_M = 0.5
L_V_DX_MINUS_M = 1.8
L_V_DY_MINUS_M = 0.4

# A leader geometry (meters), fully explicit with 2 segments:
# Segment 1: contact(midpoint) -> elbow
# Segment 2: elbow -> head(tag position)
# Horizontal segment in view:
A_H_S1_DX_M = 0.5
A_H_S1_DY_M = 0.5
A_H_S2_DX_M = 0.5
A_H_S2_DY_M = 0.5

# Vertical segment in view:
A_V_S1_DX_M = 0.5
A_V_S1_DY_M = 0.0
A_V_S2_DX_M = 0.5
A_V_S2_DY_M = 0.5

try:
    text_type = unicode
except NameError:
    text_type = str


# Offsets base are calibrated for 1:200.
SCALE_OPTIONS_ORDERED = ["1:50", "1:75", "1:100", "1:125", "1:150", "1:200", "1:500"]
SCALE_VALUES = {
    "1:50": 50,
    "1:75": 75,
    "1:100": 100,
    "1:125": 125,
    "1:150": 150,
    "1:200": 200,
    "1:500": 500,
}
# Optional correction per scale on top of (scale/200).
# Keep 1.0 for pure proportional behavior.
SCALE_CORRECTION = {
    "1:50": 1.0,
    "1:75": 1.0,
    "1:100": 1.0,
    "1:125": 1.0,
    "1:150": 1.0,
    "1:200": 1.0,
    "1:500": 1.0,
}
# Small extra shift for L tags to keep visual pairing with A in non-1:200 scales.
# Values are in meters and applied in view axes (+X right, +Y up).
L_ALIGNMENT_NUDGE_BY_SCALE_M = {
    "1:50": (0.020, 0.015),
    "1:75": (0.030, 0.020),
    "1:100": (0.040, 0.025),
    "1:125": (0.050, 0.030),
    "1:150": (0.060, 0.035),
    "1:200": (0.000, 0.000),
    "1:500": (0.000, 0.000),
}


class ScalePickerWindow(Window):
    def __init__(self, default_scale_str):
        self.Title = "Tag Segments A/L"
        self.Width = 280
        self.SizeToContent = SizeToContent.Height
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.selected = None

        root = WpfStackPanel()
        root.Margin = Thickness(16)

        lbl = WpfLabel()
        lbl.Content = "Escala de offsets (base 1:200):"
        lbl.Margin = Thickness(0, 0, 0, 6)
        root.Children.Add(lbl)

        self.listbox = WpfListBox()
        self.listbox.Margin = Thickness(0, 0, 0, 12)
        for s in SCALE_OPTIONS_ORDERED:
            item = WpfListBoxItem()
            item.Content = s
            self.listbox.Items.Add(item)

        default_index = 5  # 1:200
        if default_scale_str in SCALE_OPTIONS_ORDERED:
            default_index = SCALE_OPTIONS_ORDERED.index(default_scale_str)
        self.listbox.SelectedIndex = default_index
        root.Children.Add(self.listbox)

        btn_panel = WpfStackPanel()
        btn_panel.Orientation = WpfOrientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right

        btn_ok = WpfButton()
        btn_ok.Content = "OK"
        btn_ok.Width = 80
        btn_ok.Margin = Thickness(0, 0, 8, 0)
        btn_ok.Click += self.on_ok
        btn_panel.Children.Add(btn_ok)

        btn_cancel = WpfButton()
        btn_cancel.Content = "Cancelar"
        btn_cancel.Width = 80
        btn_cancel.Click += self.on_cancel
        btn_panel.Children.Add(btn_cancel)

        root.Children.Add(btn_panel)
        self.Content = root

    def on_ok(self, sender, args):
        if self.listbox.SelectedItem:
            self.selected = self.listbox.SelectedItem.Content
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


def get_scale_factor_from_popup(default_scale_str):
    picker = ScalePickerWindow(default_scale_str)
    result = picker.ShowDialog()
    if not result or not picker.selected:
        raise SystemExit
    return picker.selected, float(SCALE_VALUES[picker.selected]) / 200.0


def normalize_text(value):
    if not value:
        return u""
    txt = value if isinstance(value, text_type) else text_type(value)
    txt = unicodedata.normalize("NFKD", txt)
    txt = u"".join([c for c in txt if not unicodedata.combining(c)])
    return txt.lower().strip()


def get_symbol_type_name(symbol):
    p = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if p:
        return p.AsString() or ""
    return ""


def find_tag_symbol():
    symbols = list(
        FilteredElementCollector(doc)
        .OfClass(FamilySymbol)
        .OfCategory(BuiltInCategory.OST_PipeTags)
    )
    type_norm = normalize_text(TAG_TYPE_NAME)

    for v in range(VERSION_MAX, VERSION_MIN - 1, -1):
        family_name = FAMILY_BASE_NAME.format(v)
        fam_norm = normalize_text(family_name)
        for s in symbols:
            if normalize_text(s.FamilyName) != fam_norm:
                continue
            if normalize_text(get_symbol_type_name(s)) == type_norm:
                return s, family_name, v

    return None, None, None


def is_pipe_fitting(elem):
    return (
        isinstance(elem, FamilyInstance)
        and elem.Category
        and elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeFitting)
    )


def is_tee_fitting(elem):
    if not is_pipe_fitting(elem):
        return False
    try:
        mep = getattr(elem, "MEPModel", None)
        if mep and hasattr(mep, "PartType"):
            return mep.PartType == PartType.Tee
    except:
        pass
    return False


def connected_pipes_through_non_tee(pipe):
    neighbors = set()
    cm = getattr(pipe, "ConnectorManager", None)
    if not cm:
        return neighbors

    for c in cm.Connectors:
        for r in c.AllRefs:
            owner = r.Owner
            if owner is None:
                continue

            if isinstance(owner, Pipe):
                if owner.Id != pipe.Id:
                    neighbors.add(owner)
                continue

            if is_pipe_fitting(owner):
                if is_tee_fitting(owner):
                    continue
                fit_cm = getattr(getattr(owner, "MEPModel", None), "ConnectorManager", None)
                if not fit_cm:
                    continue
                for fc in fit_cm.Connectors:
                    for rr in fc.AllRefs:
                        other = rr.Owner
                        if isinstance(other, Pipe) and other.Id != pipe.Id:
                            neighbors.add(other)
    return neighbors


def build_segments(pipes):
    by_id = dict((p.Id.IntegerValue, p) for p in pipes)
    visited = set()
    segments = []

    for pid in by_id.keys():
        if pid in visited:
            continue

        stack = [by_id[pid]]
        segment = []
        while stack:
            p = stack.pop()
            pkey = p.Id.IntegerValue
            if pkey in visited:
                continue
            visited.add(pkey)
            segment.append(p)

            for n in connected_pipes_through_non_tee(p):
                nkey = n.Id.IntegerValue
                if nkey in by_id and nkey not in visited:
                    stack.append(n)

        if segment:
            segments.append(segment)

    return segments


def get_pipe_curve(pipe):
    loc = pipe.Location
    if isinstance(loc, LocationCurve):
        return loc.Curve
    return None


def get_pipe_length(pipe):
    c = get_pipe_curve(pipe)
    if c is not None:
        return c.Length
    return 0.0


def get_tipo_sistema(pipe):
    p = pipe.LookupParameter("Tipo de sistema")
    if p:
        try:
            s = p.AsString()
            if s:
                return s
        except:
            pass
        try:
            s = p.AsValueString()
            if s:
                return s
        except:
            pass
    return ""


def get_connected_pipes_from_connector(connector):
    pipes = set()
    for r in connector.AllRefs:
        owner = r.Owner
        if owner is None:
            continue

        if isinstance(owner, Pipe):
            pipes.add(owner)
            continue

        if is_pipe_fitting(owner):
            fit_cm = getattr(getattr(owner, "MEPModel", None), "ConnectorManager", None)
            if not fit_cm:
                continue
            for fc in fit_cm.Connectors:
                for rr in fc.AllRefs:
                    other = rr.Owner
                    if isinstance(other, Pipe):
                        pipes.add(other)

    return pipes


def get_tagged_pipe_ids_in_view():
    tagged = set()
    tags = (
        FilteredElementCollector(doc, view.Id)
        .OfCategory(BuiltInCategory.OST_PipeTags)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for tag in tags:
        ids = None
        try:
            ids = tag.GetTaggedLocalElementIds()
        except:
            try:
                ids = tag.GetTaggedElementIds()
            except:
                ids = None

        if not ids:
            continue

        for tid in ids:
            try:
                if tid != ElementId.InvalidElementId:
                    tagged.add(tid.IntegerValue)
            except:
                pass

    return tagged


def get_connection_point_on_pipe(pipe, other_pipe):
    cm = getattr(pipe, "ConnectorManager", None)
    if not cm:
        return None

    for c in cm.Connectors:
        for r in c.AllRefs:
            owner = r.Owner
            if owner is None:
                continue

            if isinstance(owner, Pipe) and owner.Id == other_pipe.Id:
                return c.Origin

            if is_pipe_fitting(owner) and (not is_tee_fitting(owner)):
                fit_cm = getattr(getattr(owner, "MEPModel", None), "ConnectorManager", None)
                if not fit_cm:
                    continue
                connected_to_other = False
                for fc in fit_cm.Connectors:
                    for rr in fc.AllRefs:
                        oo = rr.Owner
                        if isinstance(oo, Pipe) and oo.Id == other_pipe.Id:
                            connected_to_other = True
                            break
                    if connected_to_other:
                        break
                if connected_to_other:
                    return c.Origin

    return None


def order_segment_pipes(segment):
    by_id = dict((p.Id.IntegerValue, p) for p in segment)
    ids = set(by_id.keys())

    adjacency = {}
    for p in segment:
        pid = p.Id.IntegerValue
        adjacency[pid] = set()
        for n in connected_pipes_through_non_tee(p):
            nid = n.Id.IntegerValue
            if nid in ids:
                adjacency[pid].add(nid)

    endpoints = [pid for pid in ids if len(adjacency.get(pid, [])) <= 1]
    if endpoints:
        start = endpoints[0]
    else:
        start = next(iter(ids))

    ordered_ids = []
    visited = set()
    current = start
    prev = None

    while True:
        if current in visited:
            remaining = [pid for pid in ids if pid not in visited]
            if not remaining:
                break
            current = remaining[0]
            prev = None
            continue

        ordered_ids.append(current)
        visited.add(current)

        neighbors = [nid for nid in adjacency.get(current, []) if nid != prev and nid not in visited]
        if neighbors:
            nxt = neighbors[0]
            prev = current
            current = nxt
        else:
            remaining = [pid for pid in ids if pid not in visited]
            if not remaining:
                break
            current = remaining[0]
            prev = None

    return [by_id[pid] for pid in ordered_ids]


def get_segment_anchor_and_midpoint(segment):
    ordered = order_segment_pipes(segment)
    if not ordered:
        return None, None

    lengths = [get_pipe_length(p) for p in ordered]
    total = sum(lengths)
    if total <= 0.0:
        return None, None

    target = total * 0.5
    cumulative = 0.0

    for idx, p in enumerate(ordered):
        plen = lengths[idx]
        if plen <= 0.0:
            cumulative += plen
            continue

        if (target <= cumulative + plen) or (idx == len(ordered) - 1):
            local = (target - cumulative) / plen
            if local < 0.0:
                local = 0.0
            if local > 1.0:
                local = 1.0

            curve = get_pipe_curve(p)
            if curve is None:
                return None, None

            # Orient local parameter using previous pipe connection when possible.
            param = local
            if idx > 0:
                prev_pipe = ordered[idx - 1]
                conn_pt = get_connection_point_on_pipe(p, prev_pipe)
                if conn_pt is not None:
                    e0 = curve.GetEndPoint(0)
                    e1 = curve.GetEndPoint(1)
                    if conn_pt.DistanceTo(e1) < conn_pt.DistanceTo(e0):
                        param = 1.0 - local

            return p, curve.Evaluate(param, True)

        cumulative += plen

    return None, None


if view is None:
    raise Exception("No active view.")

view_scale = getattr(view, "Scale", 200) or 200
try:
    default_scale_str = "1:{}".format(int(view_scale))
except:
    default_scale_str = "1:200"
if default_scale_str not in SCALE_VALUES:
    default_scale_str = "1:200"

selected_scale_str, scale_factor = get_scale_factor_from_popup(default_scale_str)
scale_factor = scale_factor * SCALE_CORRECTION.get(selected_scale_str, 1.0)
# Keep A and L synchronized across scales.
leader_scale_factor = scale_factor
l_nudge_dx_m, l_nudge_dy_m = L_ALIGNMENT_NUDGE_BY_SCALE_M.get(selected_scale_str, (0.0, 0.0))
L_NUDGE_DX = l_nudge_dx_m * 3.28084
L_NUDGE_DY = l_nudge_dy_m * 3.28084
view_right = view.RightDirection
view_up = view.UpDirection

# Horizontal PLUS
A_H_DX_PLUS = A_H_DX_PLUS_M * 3.28084 * scale_factor
A_H_DY_PLUS = A_H_DY_PLUS_M * 3.28084 * scale_factor
L_H_DX_PLUS = L_H_DX_PLUS_M * 3.28084 * leader_scale_factor
L_H_DY_PLUS = L_H_DY_PLUS_M * 3.28084 * leader_scale_factor

# Horizontal MINUS
A_H_DX_MINUS = A_H_DX_MINUS_M * 3.28084 * scale_factor
A_H_DY_MINUS = A_H_DY_MINUS_M * 3.28084 * scale_factor
L_H_DX_MINUS = L_H_DX_MINUS_M * 3.28084 * leader_scale_factor
L_H_DY_MINUS = L_H_DY_MINUS_M * 3.28084 * leader_scale_factor

# Vertical PLUS
A_V_DX_PLUS = A_V_DX_PLUS_M * 3.28084 * scale_factor
A_V_DY_PLUS = A_V_DY_PLUS_M * 3.28084 * scale_factor
L_V_DX_PLUS = L_V_DX_PLUS_M * 3.28084 * leader_scale_factor
L_V_DY_PLUS = L_V_DY_PLUS_M * 3.28084 * leader_scale_factor

# Vertical MINUS
A_V_DX_MINUS = A_V_DX_MINUS_M * 3.28084 * scale_factor
A_V_DY_MINUS = A_V_DY_MINUS_M * 3.28084 * scale_factor
L_V_DX_MINUS = L_V_DX_MINUS_M * 3.28084 * leader_scale_factor
L_V_DY_MINUS = L_V_DY_MINUS_M * 3.28084 * leader_scale_factor

A_H_S1_DX = A_H_S1_DX_M * 3.28084 * leader_scale_factor
A_H_S1_DY = A_H_S1_DY_M * 3.28084 * leader_scale_factor
A_H_S2_DX = A_H_S2_DX_M * 3.28084 * leader_scale_factor
A_H_S2_DY = A_H_S2_DY_M * 3.28084 * leader_scale_factor

A_V_S1_DX = A_V_S1_DX_M * 3.28084 * leader_scale_factor
A_V_S1_DY = A_V_S1_DY_M * 3.28084 * leader_scale_factor
A_V_S2_DX = A_V_S2_DX_M * 3.28084 * leader_scale_factor
A_V_S2_DY = A_V_S2_DY_M * 3.28084 * leader_scale_factor


def offset_in_view(point, dx_right, dy_up):
    return point + view_right.Multiply(dx_right) + view_up.Multiply(dy_up)


def is_segment_horizontal_in_view(pipe):
    c = get_pipe_curve(pipe)
    if c is None:
        return True
    try:
        e0 = c.GetEndPoint(0)
        e1 = c.GetEndPoint(1)
        d = e1 - e0
        if d.GetLength() <= 1e-9:
            return True
        d = d.Normalize()
        right_comp = abs(d.DotProduct(view_right))
        up_comp = abs(d.DotProduct(view_up))
        return right_comp >= up_comp
    except:
        return True


def get_head_offset_for_pipe(pipe, tag_kind, system_text):
    """
    Return (dx_right, dy_up) from midpoint based on:
    - tag_kind: 'A' or 'L'
    - segment orientation in view: horizontal or vertical
    """
    is_h = is_segment_horizontal_in_view(pipe)
    is_minus = ("a1-" in system_text) or ("l2" in system_text) or ("-" in system_text)

    if is_h:
        if is_minus:
            if tag_kind == "A":
                return (A_H_DX_MINUS, A_H_DY_MINUS)
            return (L_H_DX_MINUS + L_NUDGE_DX, L_H_DY_MINUS + L_NUDGE_DY)
        if tag_kind == "A":
            return (A_H_DX_PLUS, A_H_DY_PLUS)
        return (L_H_DX_PLUS + L_NUDGE_DX, L_H_DY_PLUS + L_NUDGE_DY)

    if is_minus:
        if tag_kind == "A":
            return (A_V_DX_MINUS, A_V_DY_MINUS)
        return (L_V_DX_MINUS + L_NUDGE_DX, L_V_DY_MINUS + L_NUDGE_DY)
    if tag_kind == "A":
        return (A_V_DX_PLUS, A_V_DY_PLUS)
    return (L_V_DX_PLUS + L_NUDGE_DX, L_V_DY_PLUS + L_NUDGE_DY)


def force_a_tag_geometry(tag, tag_ref, anchor_pipe, midpoint, system_text):
    """Force A-tag geometry with explicit 2 leader segment lengths."""
    if is_segment_horizontal_in_view(anchor_pipe):
        # Segment 1: contact -> elbow
        elbow = offset_in_view(midpoint, A_H_S1_DX, A_H_S1_DY)
        # Segment 2: elbow -> head
        head = offset_in_view(elbow, A_H_S2_DX, A_H_S2_DY)
    else:
        # Segment 1: contact -> elbow
        elbow = offset_in_view(midpoint, A_V_S1_DX, A_V_S1_DY)
        # Segment 2: elbow -> head
        head = offset_in_view(elbow, A_V_S2_DX, A_V_S2_DY)

    try:
        tag.TagHeadPosition = head
    except:
        pass

    # Try API variants across Revit versions.
    try:
        tag.SetLeaderEnd(tag_ref, midpoint)
    except:
        try:
            tag.LeaderEnd = midpoint
        except:
            pass

    try:
        tag.SetLeaderElbow(tag_ref, elbow)
    except:
        try:
            tag.LeaderElbow = elbow
        except:
            pass


tag_symbol, family_used, version_used = find_tag_symbol()
if not tag_symbol:
    raise Exception(
        "No se encontro la etiqueta '{}' en CST_TAG Diametro Tuberia v28-v40.".format(TAG_TYPE_NAME)
    )

pipes = list(
    FilteredElementCollector(doc, view.Id)
    .OfClass(Pipe)
    .WhereElementIsNotElementType()
    .ToElements()
)
if not pipes:
    TaskDialog.Show("Info", "No hay tuberias en la vista activa.")
    raise SystemExit

segments = build_segments(pipes)
tagged_pipe_ids = get_tagged_pipe_ids_in_view()

pipe_to_segment = {}
for i, seg in enumerate(segments):
    for p in seg:
        pipe_to_segment[p.Id.IntegerValue] = i

# Exclude terminal segments directly connected to mechanical equipment
terminal_segment_ids = set()
equipos = list(
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
    .WhereElementIsNotElementType()
    .ToElements()
)
for eq in equipos:
    mepmodel = getattr(eq, "MEPModel", None)
    cm = getattr(mepmodel, "ConnectorManager", None)
    if not cm:
        continue
    for c in cm.Connectors:
        for p in get_connected_pipes_from_connector(c):
            sid = pipe_to_segment.get(p.Id.IntegerValue)
            if sid is not None:
                terminal_segment_ids.add(sid)

created = 0
skipped_terminal = 0
skipped_tagged = 0
skipped_no_midpoint = 0
skipped_no_system = 0
skipped_error = 0

t = Transaction(doc, "Tag segments between tees (A/L rules)")
t.Start()

if not tag_symbol.IsActive:
    tag_symbol.Activate()
    doc.Regenerate()

for i, seg in enumerate(segments):
    if i in terminal_segment_ids:
        skipped_terminal += 1
        continue

    seg_ids = [p.Id.IntegerValue for p in seg]
    if any((sid in tagged_pipe_ids) for sid in seg_ids):
        skipped_tagged += 1
        continue

    anchor_pipe, midpoint = get_segment_anchor_and_midpoint(seg)
    if (anchor_pipe is None) or (midpoint is None):
        skipped_no_midpoint += 1
        continue

    tipo = normalize_text(get_tipo_sistema(anchor_pipe))
    has_a = "a" in tipo
    has_l = "l" in tipo

    if not (has_a or has_l):
        skipped_no_system += 1
        continue

    # Rule: A -> with leader, L -> without leader.
    # If both appear in text, A priority applies (leader=True).
    has_leader = has_a

    try:
        tag_ref = Reference(anchor_pipe)

        tag = IndependentTag.Create(
            doc,
            view.Id,
            tag_ref,
            has_leader,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            midpoint,
        )
        tag.ChangeTypeId(tag_symbol.Id)

        # Keep tags off the segment line in view space.
        if has_leader:
            force_a_tag_geometry(tag, tag_ref, anchor_pipe, midpoint, tipo)
        else:
            dx, dy = get_head_offset_for_pipe(anchor_pipe, "L", tipo)
            try:
                tag.TagHeadPosition = offset_in_view(midpoint, dx, dy)
            except:
                pass

        created += 1
    except:
        skipped_error += 1

t.Commit()

TaskDialog.Show(
    "Resultado",
    u"Familia usada: {} (v{})\nTipo: {}\nSegmentos detectados: {}\nEtiquetas creadas: {}\nOmitidos terminales (conectados a equipo): {}\nOmitidos ya etiquetados: {}\nOmitidos sin midpoint valido: {}\nOmitidos sin A/L en 'Tipo de sistema': {}\nErrores de creacion: {}".format(
        family_used,
        version_used,
        TAG_TYPE_NAME,
        len(segments),
        created,
        skipped_terminal,
        skipped_tagged,
        skipped_no_midpoint,
        skipped_no_system,
        skipped_error,
    ),
)
