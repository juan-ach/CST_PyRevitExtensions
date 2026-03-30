# -*- coding: utf-8 -*-

__title__ = "CST Door Selector"
__author__ = "Juan Achenbach & Codex"
__version__ = "Version: 1.1"
__doc__ = """Selector visual de tipos para la familia de puertas CST_Puerta_1hoja_V31."""

import codecs
import json
import os
import re

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FamilySymbol,
    FilteredElementCollector,
    Transaction,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit

import System
from System.Drawing import Color, Font, FontStyle, Point, Size
from System.Windows.Forms import (
    Application,
    Button,
    CheckBox,
    ComboBox,
    ComboBoxStyle,
    Form,
    FormBorderStyle,
    Label,
    TextBox,
)


try:
    text_type = unicode
except NameError:
    text_type = str


doc = revit.doc
uidoc = revit.uidoc

FAMILY_NAME_PATTERN = re.compile(r"^CST_Puerta_1hoja_V(\d+)$")
MIN_FAMILY_VERSION = 31
FAMILY_DISPLAY_NAME = u"CST_Puerta_1hoja_Vx (x >= 31)"
WINDOW_TITLE = u"CST Door Selector"
STATE_FILE = os.path.join(os.path.dirname(__file__), "last_selection.json")

CLIENTE_OPTIONS = [
    (u"Dinosol", u"Dinosol"),
    (u"Bon Preu", u"Bon Preu"),
]

TIPO_OPTIONS = [
    (u"Pivotante", u"PP"),
    (u"Corredera", u"PC"),
    (u"Vaiv\u00e9n", u"PV"),
]

SERVICIO_OPTIONS = [
    (u"Positivo", u"CP"),
    (u"Negativo", u"CN"),
]

ANCHO_OPTIONS = [
    (u"0.8", u"800"),
    (u"0.9", u"900"),
    (u"1", u"1000"),
    (u"1.1", u"1100"),
    (u"1.2", u"1200"),
    (u"1.3", u"1300"),
    (u"1.4", u"1400"),
    (u"1.5", u"1500"),
    (u"1.6", u"1600"),
]

ALTO_OPTIONS = [
    (u"1.9", u"1900"),
    (u"2.0", u"2000"),
    (u"2.1", u"2100"),
]

APERTURA_OPTIONS = [
    (u"Derecha", u"ADE"),
    (u"Izquierda", u"AIZ"),
]

ESPESOR_BY_CLIENTE_SERVICIO = {
    (u"Bon Preu", u"Negativo"): u"150",
    (u"Bon Preu", u"Positivo"): u"100",
    (u"Dinosol", u"Negativo"): u"120",
    (u"Dinosol", u"Positivo"): u"80",
}

TIPO_MAP = dict(TIPO_OPTIONS)
SERVICIO_MAP = dict(SERVICIO_OPTIONS)
ANCHO_MAP = dict(ANCHO_OPTIONS)
ALTO_MAP = dict(ALTO_OPTIONS)
APERTURA_MAP = dict(APERTURA_OPTIONS)

DEFAULT_STATE = {
    "cliente": u"Dinosol",
    "servicio": u"Positivo",
    "tipo": u"Pivotante",
    "ancho": u"0.8",
    "alto": u"1.9",
    "apertura": u"Derecha",
}

OK_COLOR = Color.FromArgb(39, 174, 96)
ERROR_COLOR = Color.FromArgb(192, 57, 43)
TEXT_COLOR = Color.FromArgb(45, 45, 45)


def as_text(value):
    if value is None:
        return u""
    return text_type(value)


def load_last_selection():
    state = DEFAULT_STATE.copy()

    if not os.path.exists(STATE_FILE):
        return state

    try:
        with codecs.open(STATE_FILE, "r", "utf-8") as state_file:
            raw_data = json.load(state_file)

        for key in state.keys():
            value = raw_data.get(key)
            if value:
                state[key] = as_text(value)
    except Exception:
        pass

    return state


def save_last_selection(state):
    try:
        with codecs.open(STATE_FILE, "w", "utf-8") as state_file:
            state_file.write(json.dumps(state, ensure_ascii=False, indent=2))
    except Exception:
        pass


def get_symbol_type_name(symbol):
    param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if param:
        type_name = param.AsString()
        if type_name:
            return as_text(type_name).strip()

    name_attr = getattr(symbol, "Name", None)
    if name_attr:
        return as_text(name_attr).strip()

    return u""


def get_supported_family_version(family_name):
    """Devuelve la version soportada de la familia o None si no aplica."""
    match = FAMILY_NAME_PATTERN.match(as_text(family_name).strip())
    if not match:
        return None

    version = int(match.group(1))
    if version < MIN_FAMILY_VERSION:
        return None

    return version


def collect_door_symbols():
    symbols_by_type_name = {}

    collector = (
        FilteredElementCollector(doc)
        .OfClass(FamilySymbol)
        .OfCategory(BuiltInCategory.OST_Doors)
    )

    for symbol in collector:
        family_name = as_text(getattr(symbol, "FamilyName", u"")).strip()
        family_version = get_supported_family_version(family_name)
        if family_version is None:
            continue

        type_name = get_symbol_type_name(symbol)
        if type_name:
            # Si existen varias versiones cargadas, se prioriza la mas reciente.
            current_entry = symbols_by_type_name.get(type_name)
            if current_entry is None or family_version > current_entry["version"]:
                symbols_by_type_name[type_name] = {
                    "symbol": symbol,
                    "version": family_version,
                    "family_name": family_name,
                }

    return dict((type_name, data["symbol"]) for type_name, data in symbols_by_type_name.items())


def get_espesor_code(cliente_label, servicio_label):
    return ESPESOR_BY_CLIENTE_SERVICIO.get((cliente_label, servicio_label), u"")


def build_type_code(cliente_label, tipo_label, servicio_label, ancho_label, alto_label, apertura_label):
    if not all([cliente_label, tipo_label, servicio_label, ancho_label, alto_label, apertura_label]):
        return u""

    espesor_code = get_espesor_code(cliente_label, servicio_label)
    if not espesor_code:
        return u""

    tipologia_code = TIPO_MAP[tipo_label] + SERVICIO_MAP[servicio_label]

    return u"{}_{}_{}x{}_{}".format(
        tipologia_code,
        espesor_code,
        ANCHO_MAP[ancho_label],
        ALTO_MAP[alto_label],
        APERTURA_MAP[apertura_label],
    )


class DoorSelectorForm(Form):
    def __init__(self, symbols_by_type_name):
        Form.__init__(self)
        self.symbols_by_type_name = symbols_by_type_name
        self._updating_checkbox_group = False
        self._is_loading_defaults = True

        self.Text = WINDOW_TITLE
        self.Width = 1180
        self.Height = 305
        self.BackColor = Color.White
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ControlBox = True
        self.TopMost = True
        self.Font = Font("Segoe UI", 10)

        self._build_ui()
        self._wire_events()
        self._set_defaults()
        self._is_loading_defaults = False
        self.update_preview()

    def _build_ui(self):
        title = Label()
        title.Text = WINDOW_TITLE
        title.Location = Point(18, 16)
        title.AutoSize = True
        title.ForeColor = TEXT_COLOR
        title.Font = Font("Segoe UI", 20, FontStyle.Regular)
        self.Controls.Add(title)

        self._add_header(u"Cliente", 20, 72)
        self._add_header(u"Servicio", 155, 72)
        self._add_header(u"Tipo", 285, 72)
        self._add_header(u"Espesor", 435, 72)
        self._add_header(u"Ancho", 545, 72)
        self._add_header(u"Alto", 695, 72)
        self._add_header(u"Apertura", 845, 72)

        self.cliente_dinosol = self._add_checkbox(u"Dinosol", 22, 108, True)
        self.cliente_bon_preu = self._add_checkbox(u"Bon Preu", 22, 138, False)

        self.servicio_positivo = self._add_checkbox(u"Positivo", 157, 108, True)
        self.servicio_negativo = self._add_checkbox(u"Negativo", 157, 138, False)

        self.tipo_combo = self._add_combo([label for label, _code in TIPO_OPTIONS], 278, 104, 118)

        self.espesor_box = TextBox()
        self.espesor_box.Location = Point(428, 104)
        self.espesor_box.Size = Size(92, 28)
        self.espesor_box.ReadOnly = True
        self.espesor_box.BackColor = Color.White
        self.espesor_box.Font = Font("Consolas", 11, FontStyle.Regular)
        self.Controls.Add(self.espesor_box)

        self.ancho_combo = self._add_combo([label for label, _code in ANCHO_OPTIONS], 538, 104, 118)

        self.alto_combo = self._add_combo([label for label, _code in ALTO_OPTIONS], 688, 104, 118)

        self.apertura_derecha = self._add_checkbox(u"Derecha", 847, 108, True)
        self.apertura_izquierda = self._add_checkbox(u"Izquierda", 847, 138, False)

        self.select_button = Button()
        self.select_button.Text = u"Select"
        self.select_button.Location = Point(1010, 102)
        self.select_button.Size = Size(140, 64)
        self.select_button.Font = Font("Segoe UI", 16, FontStyle.Regular)
        self.Controls.Add(self.select_button)
        self.AcceptButton = self.select_button

        preview_label = Label()
        preview_label.Text = u"Codigo generado"
        preview_label.Location = Point(20, 196)
        preview_label.AutoSize = True
        preview_label.ForeColor = TEXT_COLOR
        preview_label.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.Controls.Add(preview_label)

        self.preview_box = TextBox()
        self.preview_box.Location = Point(20, 220)
        self.preview_box.Size = Size(810, 28)
        self.preview_box.ReadOnly = True
        self.preview_box.BackColor = Color.White
        self.preview_box.Font = Font("Consolas", 11, FontStyle.Regular)
        self.Controls.Add(self.preview_box)

        self.status_label = Label()
        self.status_label.Location = Point(850, 222)
        self.status_label.AutoSize = True
        self.status_label.ForeColor = TEXT_COLOR
        self.status_label.Font = Font("Segoe UI", 10, FontStyle.Regular)
        self.Controls.Add(self.status_label)

    def _wire_events(self):
        self.FormClosing += self.on_form_closing

        self.cliente_dinosol.CheckedChanged += self.on_cliente_changed
        self.cliente_bon_preu.CheckedChanged += self.on_cliente_changed
        self.servicio_positivo.CheckedChanged += self.on_servicio_changed
        self.servicio_negativo.CheckedChanged += self.on_servicio_changed
        self.apertura_derecha.CheckedChanged += self.on_apertura_changed
        self.apertura_izquierda.CheckedChanged += self.on_apertura_changed

        self.tipo_combo.SelectedIndexChanged += self.on_any_value_changed
        self.ancho_combo.SelectedIndexChanged += self.on_any_value_changed
        self.alto_combo.SelectedIndexChanged += self.on_any_value_changed

        self.select_button.Click += self.on_select_click

    def _set_defaults(self):
        state = load_last_selection()

        self._set_checkbox_pair(state.get("cliente"), u"Dinosol", self.cliente_dinosol, u"Bon Preu", self.cliente_bon_preu)
        self._set_checkbox_pair(state.get("servicio"), u"Positivo", self.servicio_positivo, u"Negativo", self.servicio_negativo)
        self._set_checkbox_pair(state.get("apertura"), u"Derecha", self.apertura_derecha, u"Izquierda", self.apertura_izquierda)

        self._set_combo_selection(self.tipo_combo, state.get("tipo"), 0)
        self._set_combo_selection(self.ancho_combo, state.get("ancho"), 0)
        self._set_combo_selection(self.alto_combo, state.get("alto"), 0)

    def _set_combo_selection(self, combo, target_text, fallback_index):
        target_text = target_text or u""
        match_index = combo.FindStringExact(target_text)
        if match_index >= 0:
            combo.SelectedIndex = match_index
        else:
            combo.SelectedIndex = fallback_index

    def _set_checkbox_pair(self, target_value, first_value, first_checkbox, second_value, second_checkbox):
        if target_value == second_value:
            first_checkbox.Checked = False
            second_checkbox.Checked = True
        else:
            first_checkbox.Checked = True
            second_checkbox.Checked = False

    def _add_header(self, text, x, y):
        label = Label()
        label.Text = text
        label.Location = Point(x, y)
        label.AutoSize = True
        label.ForeColor = TEXT_COLOR
        label.Font = Font("Segoe UI", 13, FontStyle.Regular)
        self.Controls.Add(label)

    def _add_checkbox(self, text, x, y, checked):
        checkbox = CheckBox()
        checkbox.Text = text
        checkbox.Location = Point(x, y)
        checkbox.AutoSize = True
        checkbox.Checked = checked
        self.Controls.Add(checkbox)
        return checkbox

    def _add_combo(self, items, x, y, width):
        combo = ComboBox()
        combo.Location = Point(x, y)
        combo.Size = Size(width, 28)
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        for item in items:
            combo.Items.Add(item)
        self.Controls.Add(combo)
        return combo

    def _selected_combo_text(self, combo):
        if combo.SelectedItem is None:
            return None
        return as_text(combo.SelectedItem)

    def _selected_cliente(self):
        if self.cliente_dinosol.Checked:
            return u"Dinosol"
        if self.cliente_bon_preu.Checked:
            return u"Bon Preu"
        return None

    def _selected_servicio(self):
        if self.servicio_positivo.Checked:
            return u"Positivo"
        if self.servicio_negativo.Checked:
            return u"Negativo"
        return None

    def _selected_apertura(self):
        if self.apertura_derecha.Checked:
            return u"Derecha"
        if self.apertura_izquierda.Checked:
            return u"Izquierda"
        return None

    def _current_state(self):
        return {
            "cliente": self._selected_cliente() or DEFAULT_STATE["cliente"],
            "servicio": self._selected_servicio() or DEFAULT_STATE["servicio"],
            "tipo": self._selected_combo_text(self.tipo_combo) or DEFAULT_STATE["tipo"],
            "ancho": self._selected_combo_text(self.ancho_combo) or DEFAULT_STATE["ancho"],
            "alto": self._selected_combo_text(self.alto_combo) or DEFAULT_STATE["alto"],
            "apertura": self._selected_apertura() or DEFAULT_STATE["apertura"],
        }

    def _current_code(self):
        return build_type_code(
            self._selected_cliente(),
            self._selected_combo_text(self.tipo_combo),
            self._selected_servicio(),
            self._selected_combo_text(self.ancho_combo),
            self._selected_combo_text(self.alto_combo),
            self._selected_apertura(),
        )

    def _apply_exclusive_checkbox_logic(self, current_checkbox, other_checkbox):
        if self._updating_checkbox_group:
            return

        self._updating_checkbox_group = True
        try:
            if current_checkbox.Checked:
                other_checkbox.Checked = False
            elif not other_checkbox.Checked:
                current_checkbox.Checked = True
        finally:
            self._updating_checkbox_group = False

        self.update_preview()

    def on_cliente_changed(self, sender, args):
        if sender == self.cliente_dinosol:
            self._apply_exclusive_checkbox_logic(self.cliente_dinosol, self.cliente_bon_preu)
        else:
            self._apply_exclusive_checkbox_logic(self.cliente_bon_preu, self.cliente_dinosol)

    def on_servicio_changed(self, sender, args):
        if sender == self.servicio_positivo:
            self._apply_exclusive_checkbox_logic(self.servicio_positivo, self.servicio_negativo)
        else:
            self._apply_exclusive_checkbox_logic(self.servicio_negativo, self.servicio_positivo)

    def on_apertura_changed(self, sender, args):
        if sender == self.apertura_derecha:
            self._apply_exclusive_checkbox_logic(self.apertura_derecha, self.apertura_izquierda)
        else:
            self._apply_exclusive_checkbox_logic(self.apertura_izquierda, self.apertura_derecha)

    def on_any_value_changed(self, sender, args):
        self.update_preview()

    def on_form_closing(self, sender, args):
        save_last_selection(self._current_state())

    def update_preview(self):
        espesor_code = get_espesor_code(self._selected_cliente(), self._selected_servicio())
        self.espesor_box.Text = espesor_code

        code = self._current_code()
        self.preview_box.Text = code

        if not self._is_loading_defaults:
            save_last_selection(self._current_state())

        if not code:
            self.status_label.Text = u"Faltan opciones"
            self.status_label.ForeColor = ERROR_COLOR
            return

        if code in self.symbols_by_type_name:
            self.status_label.Text = u"Tipo disponible"
            self.status_label.ForeColor = OK_COLOR
        else:
            self.status_label.Text = u"Tipo no encontrado"
            self.status_label.ForeColor = ERROR_COLOR

    def on_select_click(self, sender, args):
        type_code = self._current_code()

        if not type_code:
            TaskDialog.Show(WINDOW_TITLE, u"Completa todas las opciones antes de continuar.")
            return

        symbol = self.symbols_by_type_name.get(type_code)
        if symbol is None:
            TaskDialog.Show(
                WINDOW_TITLE,
                u"No existe el tipo '{}'\nen una familia compatible '{}'.".format(type_code, FAMILY_DISPLAY_NAME),
            )
            return

        tx = None
        try:
            if not symbol.IsActive:
                tx = Transaction(doc, u"Activar tipo de puerta")
                tx.Start()
                symbol.Activate()
                doc.Regenerate()
                tx.Commit()
                tx = None

            save_last_selection(self._current_state())
            uidoc.PostRequestForElementTypePlacement(symbol)
            self.Close()
        except Exception as ex:
            if tx is not None:
                try:
                    tx.RollBack()
                except Exception:
                    pass

            TaskDialog.Show(
                WINDOW_TITLE,
                u"No se pudo dejar preparado el tipo '{}'.\n\n{}".format(type_code, as_text(ex)),
            )


def main():
    if doc is None or uidoc is None:
        TaskDialog.Show(WINDOW_TITLE, u"No hay un documento de Revit activo.")
        return

    if doc.IsFamilyDocument:
        TaskDialog.Show(WINDOW_TITLE, u"Este comando debe ejecutarse en un proyecto, no dentro del editor de familias.")
        return

    symbols_by_type_name = collect_door_symbols()
    if not symbols_by_type_name:
        TaskDialog.Show(
            WINDOW_TITLE,
            u"No se encontraron tipos cargados de familias compatibles '{}' en este proyecto.".format(FAMILY_DISPLAY_NAME),
        )
        return

    Application.EnableVisualStyles()
    form = DoorSelectorForm(symbols_by_type_name)
    form.ShowDialog()


if __name__ == "__main__":
    main()
