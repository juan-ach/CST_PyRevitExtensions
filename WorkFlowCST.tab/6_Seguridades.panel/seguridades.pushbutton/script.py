# -*- coding: utf-8 -*-
__title__ = "Safety Tags"
__author__ = "Juan + ChatGPT"

from pyrevit import revit, script
from Autodesk.Revit.DB import *

doc = revit.doc
view = doc.ActiveView
uidoc = revit.uidoc

# --- CONFIGURATION ---
TAG_FAMILY_NAMES = ["CST_TAG Seguridad_v3", "CST_TAG Seguridad_v3 - cat"]
TYPE_SMALL = "Detalle 3" # < 20,000
TYPE_LARGE = "Detalle 2" # >= 20,000
TYPE_CONG = "Detalle 1"  # Override
TYPE_O = "Detalle 4"     # Override
THRESHOLD_VOL = 20000.0

EXCLUDED_KEYWORDS = ["Cong", "Arcon", "Isla", "CAB", "Mural", "Semi", "Vitrina"]

# --- HELPERS ---

def _norm(s):
    """Normalize string for comparison."""
    if not s: return ""
    return s.strip().lower()

def is_excluded(eq):
    """Check if equipment Type Name contains excluded keywords."""
    try:
        # Get Type
        el_type = doc.GetElement(eq.GetTypeId())
        if not el_type: return False
        
        # Get Type Name
        # Try BuiltInParameter first
        p = el_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name = p.AsString() if p else el_type.Name
        
        if not type_name: return False
        
        # Check keywords
        for kw in EXCLUDED_KEYWORDS:
            if kw.lower() in type_name.lower():
                return True
    except:
        pass
    return False

def get_ubicacion(eq):
    """Get 'ubicación' parameter value."""
    p = eq.LookupParameter("ubicación")
    if p:
        return p.AsString()
    return None

def get_room_volume_value(room):
    """
    Get volume value. 
    Tries 'Volumen' (User Param) then 'Volume' (BuiltIn).
    Returns float (Liters logic if string contains 'L').
    """
    val = 0.0
    
    # 1. Try 'Volumen' (Explicit user param, likely Liters)
    p = room.LookupParameter("Volumen")
    if not p:
        # 2. Try 'Volume' (BuiltIn, likely Cubic Feet)
        p = room.get_Parameter(BuiltInParameter.ROOM_VOLUME)
    
    if p and p.HasValue:
        # String parsing (e.g. "25000 L")
        if p.StorageType == StorageType.String:
            s_val = p.AsString()
            # Clean units
            clean = s_val.lower().replace("l", "").replace("m³", "").strip()
            try:
                val = float(clean)
            except:
                val = 0.0
        # Double parsing
        elif p.StorageType == StorageType.Double:
            val = p.AsDouble()
            # Note: If it came from BuiltInParameter ROOM_VOLUME, it is in Cubic Feet!
            # If came from custom 'Volumen' as Double, assuming it is already correct unit (Liters?)
            # Heuristic: If value is small (< 1000), it might be m3 or ft3. If > 1000, likely Liters.
            # But user logic creates a hard threshold of 20000. 
            # 20 m3 = 20000 L. 
            # 20 ft3 = 566 L. 
            # If we get raw ft3 (e.g. 700 ft3 = 20000L).
            # Let's assume custom param 'Volumen' is Liters (Double) or String.
            # If we fall back to BuiltIn ROOM_VOLUME, we convert ft3 to Liters.
            if p.Definition.BuiltInParameter == BuiltInParameter.ROOM_VOLUME:
                val = val * 28.3168 # ft3 to Liters
            
    return val

def get_center(eq):
    """Get center point for tag."""
    loc = eq.Location
    if isinstance(loc, LocationPoint):
        return loc.Point
    return None

# --- MAIN ---

# 1. Collect Rooms and map Name -> Room
rooms_map = {}
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()


def get_room_name_safe(element):
    """Safely get room name."""
    # 1. BuiltInParameter (Most reliable)
    try:
        p = element.get_Parameter(BuiltInParameter.ROOM_NAME)
        if p and p.HasValue:
            return p.AsString()
    except:
        pass
        
    # 2. Custom Param 'Nombre'
    try:
        p = element.LookupParameter("Nombre")
        if p and p.HasValue:
            return p.AsString()
    except:
        pass
        
    # 3. .Name Property (Can fail)
    try:
        return element.Name
    except:
        pass
        
    return None

for r in rooms:
    r_name = get_room_name_safe(r)
    if r_name:
        rooms_map[_norm(r_name)] = r

# 2. Collect Equipment
equips = FilteredElementCollector(doc, view.Id).\
         OfCategory(BuiltInCategory.OST_MechanicalEquipment).\
         WhereElementIsNotElementType().ToElements()

# 3. Get Tag Symbols
fam_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_MechanicalEquipmentTags).ToElements()
tag_sym_small = None # Detalle 3
tag_sym_large = None # Detalle 2
tag_sym_cong = None  # Detalle 1
tag_sym_o = None     # Detalle 4

# Identify which family is present
found_family_name = None
present_families = set()

for fs in fam_symbols:
    try:
        fam_name = fs.FamilyName
    except:
        fam_name = fs.Family.Name
    present_families.add(fam_name)

# Pick the first one that exists
for name in TAG_FAMILY_NAMES:
    if name in present_families:
        found_family_name = name
        break

if not found_family_name:
    script.get_output().print_md("ERROR: Could not find any of the required Tag Families: {}".format(", ".join(TAG_FAMILY_NAMES)))
    script.exit()

# Now extract symbols for that family
for fs in fam_symbols:
    try:
        fam_name = fs.FamilyName
    except:
        fam_name = fs.Family.Name
        
    if fam_name != found_family_name: continue
    
    # Check Type Name
    p_sym = fs.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if p_sym:
        type_name = p_sym.AsString()
        
        if type_name == TYPE_SMALL: tag_sym_small = fs
        elif type_name == TYPE_LARGE: tag_sym_large = fs
        elif type_name == TYPE_CONG: tag_sym_cong = fs
        elif type_name == TYPE_O: tag_sym_o = fs

if not all([tag_sym_small, tag_sym_large, tag_sym_cong, tag_sym_o]):
    script.get_output().print_md("ERROR: Missing one or more Tag Types in Family '{}'".format(found_family_name))
    script.exit()

# 4. Process
t = Transaction(doc, "Tag Equipment by Volume")
t.Start()

# Activate symbols
if not tag_sym_small.IsActive: tag_sym_small.Activate()
if not tag_sym_large.IsActive: tag_sym_large.Activate()
if not tag_sym_cong.IsActive: tag_sym_cong.Activate()
if not tag_sym_o.IsActive: tag_sym_o.Activate()


count_tagged = 0

for eq in equips:
    # A. Check Exclusions
    if is_excluded(eq):
        continue
        
    # B. Get Ubicacion
    ubic = get_ubicacion(eq)
    if not ubic:
        continue
        
    # C. Decide Tag (Override vs Volume)
    target_sym = None
    u_upper = ubic.upper()
    
    # 1. Overrides
    if "CONG" in u_upper:
        target_sym = tag_sym_cong
    elif "O." in u_upper:
        target_sym = tag_sym_o
    
    # 2. Volume Logic
    if not target_sym:
        # Require Room Match
        room = rooms_map.get(_norm(ubic))
        if not room:
            continue
            
        vol = get_room_volume_value(room)
        
        if vol < THRESHOLD_VOL:
            target_sym = tag_sym_small
        else:
            target_sym = tag_sym_large
        
    # F. Place Tag

    pt = get_center(eq)
    if pt:
        # Offset 1 inch in Y (1/12 feet)
        pt = XYZ(pt.X, pt.Y + 2.5, pt.Z)

        try:
            # Create Independent Tag
            # Signature: Create(Document, ElementId (View), Reference, bool (Leader), TagMode, TagOrientation, XYZ)
            # OR Revit 2018+: Create(Document, ElementId (TagType), ElementId (OwnerView), Reference, bool (Leader), TagOrientation, XYZ)
            
            # Using 2018+ signature usually safer for newer PyRevit, but let's try standard approach ensuring types match
            # We can use IndependentTag.Create generic and then ChangeTypeId
            
            # Create tag
            new_tag = IndependentTag.Create(doc, view.Id, Reference(eq), False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, pt)
            new_tag.ChangeTypeId(target_sym.Id)
            
            count_tagged += 1
        except Exception as e:
            pass

t.Commit()

# Report
print("Tagged {} elements.".format(count_tagged))