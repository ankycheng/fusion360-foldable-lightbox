import adsk.core
import adsk.fusion
import json
import math
import os
import traceback

_app = None
_ui = None
_handlers = []

CMD_ID = "FoldableLightbox_cmd"
CMD_NAME = "Foldable Lightbox"
CMD_DESC = "Parametric foldable lightbox (Triangle / Trapezoid / Square) with optional end caps"

_SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_settings.json")


def _load_settings():
    try:
        with open(_SETTINGS_PATH, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_settings(d):
    try:
        with open(_SETTINGS_PATH, "w") as f:
            json.dump(d, f, indent=2)
    except Exception as ex:
        _log(f"save settings failed: {ex}")


def _vi(s, key, default_str):
    """ValueInput helper: saved cm value if present, else fallback string default."""
    if key in s and isinstance(s[key], (int, float)):
        return adsk.core.ValueInput.createByReal(float(s[key]))
    return adsk.core.ValueInput.createByString(default_str)


def run(context):
    global _app, _ui
    try:
        _app = adsk.core.Application.get()
        _ui = _app.userInterface

        cmd_def = _ui.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            cmd_def.deleteMe()
        cmd_def = _ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESC)

        on_created = CommandCreatedHandler()
        cmd_def.commandCreated.add(on_created)
        _handlers.append(on_created)

        ws = _ui.workspaces.itemById("FusionSolidEnvironment")
        panel = ws.toolbarPanels.itemById("SolidCreatePanel")
        if not panel.controls.itemById(CMD_ID):
            panel.controls.addCommand(cmd_def)
    except Exception:
        if _ui:
            _ui.messageBox(f"Foldable Lightbox failed to start:\n{traceback.format_exc()}")


def stop(context):
    global _handlers
    try:
        ws = _ui.workspaces.itemById("FusionSolidEnvironment")
        panel = ws.toolbarPanels.itemById("SolidCreatePanel")
        ctrl = panel.controls.itemById(CMD_ID)
        if ctrl:
            ctrl.deleteMe()
        cmd_def = _ui.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            cmd_def.deleteMe()
    except Exception:
        pass
    _handlers = []


def _log(msg):
    try:
        palette = _ui.palettes.itemById("TextCommands")
        if palette:
            palette.writeText(f"[FoldableLightbox] {msg}")
    except Exception:
        pass


class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        try:
            cmd = args.command
            cmd.isExecutedWhenPreEmpted = False
            inputs = cmd.commandInputs
            s = _load_settings()

            prof = inputs.addDropDownCommandInput(
                "profile", "Profile", adsk.core.DropDownStyles.TextListDropDownStyle
            )
            sel_profile = s.get("profile", "Trapezoid")
            for name in ("Triangle", "Trapezoid", "Square"):
                prof.listItems.add(name, name == sel_profile)

            g_dim = inputs.addGroupCommandInput("g_dim", "Dimensions")
            g_dim.children.addValueInput("top_w", "Top Width", "mm",
                _vi(s, "top_w", "18 mm"))
            g_dim.children.addValueInput("bot_w", "Bottom Width", "mm",
                _vi(s, "bot_w", "32 mm"))
            g_dim.children.addValueInput("height", "Profile Height", "mm",
                _vi(s, "height", "25 mm"))
            g_dim.children.addValueInput("depth", "Depth (axial length)", "mm",
                _vi(s, "depth", "45 mm"))

            g_sheet = inputs.addGroupCommandInput("g_sheet", "Sheet & Hinge")
            g_sheet.children.addValueInput("sheet_t", "Sheet Thickness", "mm",
                _vi(s, "sheet_t", "1.6 mm"))
            g_sheet.children.addValueInput("hinge_t", "Living Hinge Thickness", "mm",
                _vi(s, "hinge_t", "0.4 mm"))
            g_sheet.children.addValueInput("hinge_w", "Hinge Groove Width", "mm",
                _vi(s, "hinge_w", "0.8 mm"))

            g_text = inputs.addGroupCommandInput("g_text", "Text (separate body for multi-material)")
            g_text.children.addStringValueInput("text_str", "Text (Front)", s.get("text_str", "TAXI"))
            g_text.children.addStringValueInput("text_str_back",
                "Back Text (empty = same as front)", s.get("text_str_back", ""))

            font_dd = g_text.children.addDropDownCommandInput(
                "text_font", "Font", adsk.core.DropDownStyles.TextListDropDownStyle
            )
            _FONTS = [
                "Impact", "Arial Black", "Helvetica Neue", "Arial", "Futura",
                "Avenir Next", "Verdana", "Tahoma", "Gill Sans", "Optima",
                "Georgia", "Times New Roman", "Didot", "Courier New", "Menlo",
                "Custom...",
            ]
            sel_font = s.get("text_font_ui", "Impact")
            if sel_font not in _FONTS:
                sel_font = "Impact"
            for f in _FONTS:
                font_dd.listItems.add(f, f == sel_font)

            g_text.children.addStringValueInput("text_font_custom",
                "Custom Font (used when 'Custom...' selected)", s.get("text_font_custom", ""))
            g_text.children.addBoolValueInput("text_bold", "Bold", True, "", s.get("text_bold", True))
            g_text.children.addBoolValueInput("text_italic", "Italic", True, "", s.get("text_italic", False))
            g_text.children.addValueInput("text_h", "Text Height", "mm",
                _vi(s, "text_h", "12 mm"))
            g_text.children.addValueInput("text_extrude", "Text Body Depth (mode-dependent)", "mm",
                _vi(s, "text_extrude", "0.6 mm"))
            mode_dd = g_text.children.addDropDownCommandInput(
                "text_mode", "Text Mode", adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            _TEXT_MODES = ["Flush Recess", "Through-cut", "Emboss", "Through-cut + Emboss"]
            sel_mode = s.get("text_mode")
            if sel_mode not in _TEXT_MODES:
                # Migrate legacy text_through bool if present
                sel_mode = "Through-cut" if s.get("text_through") else "Flush Recess"
            for m in _TEXT_MODES:
                mode_dd.listItems.add(m, m == sel_mode)
            g_text.children.addBoolValueInput("text_autofit", "Auto-fit Text to Panel", True, "",
                s.get("text_autofit", True))
            g_text.children.addBoolValueInput("text_autosize_depth",
                "Auto-size Depth to Fit Text", True, "", s.get("text_autosize_depth", False))

            g_cap = inputs.addGroupCommandInput("g_cap", "End Caps (with recess for sheet)")
            g_cap.children.addBoolValueInput("endcaps", "Generate End Caps", True, "",
                s.get("endcaps", True))
            g_cap.children.addValueInput("endcap_t", "Cap Plate Thickness", "mm",
                _vi(s, "endcap_t", "4 mm"))
            g_cap.children.addValueInput("endcap_ring_w", "Outer Ring Thickness (printable, >= 0.8mm)", "mm",
                _vi(s, "endcap_ring_w", "0.8 mm"))
            g_cap.children.addValueInput("endcap_recess", "Recess Depth (into inner face)", "mm",
                _vi(s, "endcap_recess", "2.5 mm"))
            g_cap.children.addValueInput("endcap_clr", "Recess Clearance (per side)", "mm",
                _vi(s, "endcap_clr", "0.4 mm"))
            g_cap.children.addValueInput("endcap_corner_r", "Outer Corner Fillet Radius (mm, 0 = sharp)", "mm",
                _vi(s, "endcap_corner_r", "1.5 mm"))

            g_appr = inputs.addGroupCommandInput("g_appr", "Appearance (color / material)")
            g_appr.children.addBoolValueInput("auto_appearance", "Auto-apply Appearance", True, "",
                s.get("auto_appearance", True))

            _COLORS = ["White", "Black", "Yellow", "Red", "Green", "Translucent", "None"]
            body_dd = g_appr.children.addDropDownCommandInput(
                "body_color", "Body Color (Sheet + EndCap)",
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            sel_body = s.get("body_color", "Yellow")
            if sel_body not in _COLORS:
                sel_body = "Yellow"
            for c in _COLORS:
                body_dd.listItems.add(c, c == sel_body)

            text_dd = g_appr.children.addDropDownCommandInput(
                "text_color", "Text Color",
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            _TEXT_COLORS = ["Black", "White", "Red", "Yellow", "Green", "None"]
            sel_text_color = s.get("text_color", "Black")
            if sel_text_color not in _TEXT_COLORS:
                sel_text_color = "Black"
            for c in _TEXT_COLORS:
                text_dd.listItems.add(c, c == sel_text_color)

            on_exec = CommandExecuteHandler()
            cmd.execute.add(on_exec)
            _handlers.append(on_exec)

            on_val = ValidateInputsHandler()
            cmd.validateInputs.add(on_val)
            _handlers.append(on_val)

            on_changed = InputChangedHandler()
            cmd.inputChanged.add(on_changed)
            _handlers.append(on_changed)

            _apply_profile_visibility(inputs, sel_profile)
        except Exception:
            if _ui:
                _ui.messageBox(f"CommandCreated failed:\n{traceback.format_exc()}")


def _apply_profile_visibility(inputs, profile):
    """Enable/disable width inputs based on the selected profile:
    - Triangle: top_w unused (only bot_w + height matter)
    - Square:   bot_w is forced = top_w, so disable bot_w and keep it synced
    - Trapezoid: both enabled"""
    top_w = _item(inputs, "top_w")
    bot_w = _item(inputs, "bot_w")
    if not top_w or not bot_w:
        return
    if profile == "Triangle":
        top_w.isEnabled = False
        bot_w.isEnabled = True
    elif profile == "Square":
        top_w.isEnabled = True
        bot_w.isEnabled = False
        try:
            bot_w.value = top_w.value
        except Exception:
            pass
    else:  # Trapezoid
        top_w.isEnabled = True
        bot_w.isEnabled = True


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args):
        try:
            inputs = args.inputs
            changed = args.input
            prof_inp = _item(inputs, "profile")
            if not prof_inp or not prof_inp.selectedItem:
                return
            profile = prof_inp.selectedItem.name
            if changed.id == "profile":
                _apply_profile_visibility(inputs, profile)
            elif changed.id == "top_w" and profile == "Square":
                bot_w = _item(inputs, "bot_w")
                if bot_w:
                    try:
                        bot_w.value = changed.value
                    except Exception:
                        pass
        except Exception:
            pass


class ValidateInputsHandler(adsk.core.ValidateInputsEventHandler):
    def notify(self, args):
        try:
            inp = args.inputs
            profile = _item(inp, "profile").selectedItem.name
            top = _item(inp, "top_w").value
            bot = _item(inp, "bot_w").value
            h = _item(inp, "height").value
            d = _item(inp, "depth").value
            t = _item(inp, "sheet_t").value
            ht = _item(inp, "hinge_t").value
            text_h = _item(inp, "text_h").value
            text_e = _item(inp, "text_extrude").value
            tm_inp = _item(inp, "text_mode")
            text_mode = tm_inp.selectedItem.name if tm_inp and tm_inp.selectedItem else "Flush Recess"
            cap_t = _item(inp, "endcap_t").value
            cap_r = _item(inp, "endcap_recess").value
            ring_w = _item(inp, "endcap_ring_w").value
            cap_c = _item(inp, "endcap_clr").value
            corner_r = _item(inp, "endcap_corner_r").value

            # text_extrude constraints vary by mode:
            #   Flush Recess: 0 < text_e < t  (partial recess)
            #   Through-cut: ignored (forced = t)
            #   Emboss / Through-cut + Emboss: text_e > 0 (emboss height above sheet)
            if text_mode == "Through-cut":
                te_ok = True
            elif text_mode in ("Emboss", "Through-cut + Emboss"):
                te_ok = text_e > 0
            else:  # Flush Recess
                te_ok = 0 < text_e < t

            ok = (h > 0.05 and d > 0.1 and t > 0.01 and ht > 0 and ht < t
                  and top > 0 and bot > 0 and text_h >= 0
                  and te_ok and cap_r < cap_t
                  and ring_w > 0 and cap_c >= 0 and ring_w + cap_c < t
                  and corner_r >= 0)
            if profile == "Trapezoid":
                ok = ok and top <= bot
            args.areInputsValid = bool(ok)
        except Exception:
            args.areInputsValid = False


def _item(inputs, cid):
    """Find a command input by id, searching nested groups if needed."""
    it = inputs.itemById(cid)
    if it:
        return it
    for i in range(inputs.count):
        child = inputs.item(i)
        if child.objectType == adsk.core.GroupCommandInput.classType():
            found = child.children.itemById(cid)
            if found:
                return found
    return None


def _resolve_font(inp):
    """Return font name from dropdown; if 'Custom...' is selected, read the text field."""
    dd = _item(inp, "text_font")
    sel = dd.selectedItem.name if dd else "Arial"
    if sel == "Custom...":
        custom = _item(inp, "text_font_custom")
        return (custom.value.strip() if custom and custom.value.strip() else "Arial")
    return sel


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        try:
            inp = args.command.commandInputs
            profile = _item(inp, "profile").selectedItem.name
            top_w = _item(inp, "top_w").value
            bot_w = _item(inp, "bot_w").value
            if profile == "Square":
                bot_w = top_w  # force equal for square

            p = {
                "profile":       profile,
                "top_w":         top_w,
                "bot_w":         bot_w,
                "height":        _item(inp, "height").value,
                "depth":         _item(inp, "depth").value,
                "sheet_t":       _item(inp, "sheet_t").value,
                "hinge_t":       _item(inp, "hinge_t").value,
                "hinge_w":       _item(inp, "hinge_w").value,
                "text_str":      _item(inp, "text_str").value or "",
                "text_str_back": _item(inp, "text_str_back").value or "",
                "text_font":     _resolve_font(inp),
                "text_bold":     _item(inp, "text_bold").value,
                "text_italic":   _item(inp, "text_italic").value,
                "text_h":        _item(inp, "text_h").value,
                "text_extrude":  _item(inp, "text_extrude").value,
                "text_mode":     _item(inp, "text_mode").selectedItem.name,
                "text_autofit":  _item(inp, "text_autofit").value,
                "text_autosize_depth": _item(inp, "text_autosize_depth").value,
                "endcaps":       _item(inp, "endcaps").value,
                "endcap_t":      _item(inp, "endcap_t").value,
                "endcap_ring_w": _item(inp, "endcap_ring_w").value,
                "endcap_recess": _item(inp, "endcap_recess").value,
                "endcap_clr":    _item(inp, "endcap_clr").value,
                "endcap_corner_r": _item(inp, "endcap_corner_r").value,
                "auto_appearance": _item(inp, "auto_appearance").value,
                "body_color":      _item(inp, "body_color").selectedItem.name,
                "text_color":      _item(inp, "text_color").selectedItem.name,
            }
            # Persist UI state for next invocation. Save BEFORE build so a failed
            # build still remembers what the user tried (they can reopen and tweak).
            save_dict = dict(p)
            save_dict["text_font_ui"] = _item(inp, "text_font").selectedItem.name
            save_dict["text_font_custom"] = _item(inp, "text_font_custom").value or ""
            _save_settings(save_dict)

            design = adsk.fusion.Design.cast(_app.activeProduct)
            if not design:
                _ui.messageBox("Please create or open a design first.")
                return
            build_lightbox(design, p)
        except Exception:
            if _ui:
                _ui.messageBox(f"Execute failed:\n{traceback.format_exc()}")


# ============================================================================
# Geometry builders
# ============================================================================

def compute_panels(profile, top_w, bot_w, height):
    """Return [(name, panel_length_along_Y)] from bottom (base) upward in flat layout."""
    if profile == "Triangle":
        slant = math.sqrt(height ** 2 + (bot_w / 2.0) ** 2)
        return [("base", bot_w), ("front", slant), ("back", slant)]
    if profile == "Square":
        return [("base", bot_w), ("front", height), ("top", top_w), ("back", height)]
    if profile == "Trapezoid":
        slant = math.sqrt(height ** 2 + ((bot_w - top_w) / 2.0) ** 2)
        return [("base", bot_w), ("front", slant), ("top", top_w), ("back", slant)]
    raise ValueError(f"Unknown profile: {profile}")


def cumulative_panel_ranges(panels):
    out = {}
    y = 0.0
    for name, w in panels:
        out[name] = (y, y + w, y + w / 2.0)
        y += w
    return out


def build_lightbox(design, p):
    root = design.rootComponent

    parent_occ = root.occurrences.addNewComponent(adsk.core.Matrix3D.create())
    parent = parent_occ.component
    parent.name = f"Lightbox_{p['profile']}"

    sheet_occ = parent.occurrences.addNewComponent(adsk.core.Matrix3D.create())
    sheet_comp = sheet_occ.component
    sheet_comp.name = "FlatSheet"

    panels = compute_panels(p["profile"], p["top_w"], p["bot_w"], p["height"])
    depth = p["depth"]
    t = p["sheet_t"]

    # If auto-size-depth is on, measure the text at its requested height+font and
    # compute the minimum depth needed, overriding the user-supplied depth.
    if p.get("text_autosize_depth") and p.get("text_str") and p["text_h"] > 0:
        tmp_plane = _offset_plane(sheet_comp, sheet_comp.xYConstructionPlane, t, "__tmp_autosize")
        depth = _autosize_depth(sheet_comp, tmp_plane, p)
        try:
            tmp_plane.deleteMe()
        except Exception:
            pass
        # Autosize can shrink depth below what the end caps need. thin_sheet_ends
        # silently skips when recess_d >= depth/2, but build_end_caps still creates
        # caps with a rabbet pocket — the parts won't mate. Bump depth to guarantee
        # a thinned band + middle region + thinned band fit.
        if p["endcaps"]:
            recess_d = p["endcap_recess"]
            middle_min = 0.2  # 2 mm middle of sheet between the two rabbet bands
            need = 2 * recess_d + middle_min
            if depth < need:
                _log(f"autosize: depth {depth*10:.1f}mm bumped to {need*10:.1f}mm to clear end-cap recess bands")
                depth = need

    panel_ranges = cumulative_panel_ranges(panels)

    _log(f"profile={p['profile']} panels={[(n, round(w, 3)) for n, w in panels]} depth={depth}")

    sheet_body = build_flat_sheet(sheet_comp, panels, depth, t)

    top_plane = _offset_plane(sheet_comp, sheet_comp.xYConstructionPlane, t, "top_sketch_plane")
    add_hinge_grooves(sheet_comp, top_plane, panel_ranges, depth, t, p)

    if p["text_str"] and p["text_h"] > 0:
        add_text_bodies(sheet_comp, top_plane, panel_ranges, depth, t, p)

    if p["endcaps"]:
        total_len = sum(w for _, w in panels)
        thin_depth = p["endcap_ring_w"] + p["endcap_clr"]
        thin_sheet_ends(sheet_comp, top_plane, depth, total_len, t,
                        thin_depth, p["endcap_recess"])
        build_end_caps(parent, panels, depth, t, p)

    if p.get("auto_appearance"):
        apply_appearances(design, parent, p["body_color"], p["text_color"])


_APPEARANCE_MAP = {
    # name → (preferred Fusion library appearance name, RGBA fallback)
    "White":       ("Plastic - Matte (White)",             (245, 245, 245, 255)),
    "Black":       ("Plastic - Matte (Black)",             (30, 30, 30, 255)),
    "Yellow":      ("Plastic - Matte (Yellow)",            (255, 200, 0, 255)),
    "Red":         ("Plastic - Matte (Red)",               (200, 30, 30, 255)),
    "Green":       ("Plastic - Matte (Green)",             (40, 170, 70, 255)),
    "Translucent": ("Plastic - Translucent Matte (White)", (230, 230, 230, 120)),
}


def _find_library_appearance(name):
    """Look up an appearance by name across Fusion's material libraries."""
    app = adsk.core.Application.get()
    for lib_name in ("Fusion Appearance Library", "Fusion Material Library"):
        lib = app.materialLibraries.itemByName(lib_name)
        if not lib:
            continue
        appr = lib.appearances.itemByName(name)
        if appr:
            return appr
    return None


def _get_or_create_appearance(design, color_name):
    """Return a usable Appearance for the given color label, or None if 'None'.
    Tries: design cache → Fusion library copy → RGB fallback (copy any plastic + recolor)."""
    if not color_name or color_name == "None":
        return None
    spec = _APPEARANCE_MAP.get(color_name)
    if not spec:
        return None
    lib_appr_name, rgba = spec

    existing = design.appearances.itemByName(lib_appr_name)
    if existing:
        return existing

    lib_appr = _find_library_appearance(lib_appr_name)
    if lib_appr:
        try:
            return design.appearances.addByCopy(lib_appr, lib_appr_name)
        except Exception as ex:
            _log(f"appearance copy '{lib_appr_name}' failed: {ex}")

    # RGB fallback: copy any plastic appearance we can find and override its color.
    custom_name = f"Custom_{color_name}"
    cached = design.appearances.itemByName(custom_name)
    if cached:
        return cached

    app = adsk.core.Application.get()
    template = None
    for lib_name in ("Fusion Appearance Library", "Fusion Material Library"):
        lib = app.materialLibraries.itemByName(lib_name)
        if not lib:
            continue
        for i in range(lib.appearances.count):
            a = lib.appearances.item(i)
            if "plastic" in a.name.lower():
                template = a
                break
        if template:
            break
    if not template:
        _log(f"appearance '{color_name}' fallback: no plastic template found")
        return None

    try:
        new_appr = design.appearances.addByCopy(template, custom_name)
        for i in range(new_appr.appearanceProperties.count):
            prop = new_appr.appearanceProperties.item(i)
            if prop.objectType == adsk.core.ColorProperty.classType():
                prop.value = adsk.core.Color.create(rgba[0], rgba[1], rgba[2], rgba[3])
                break
        return new_appr
    except Exception as ex:
        _log(f"appearance custom '{custom_name}' failed: {ex}")
        return None


def apply_appearances(design, parent_comp, body_color, text_color):
    """Assign appearances by PARENT COMPONENT, not body name. Any body inside a
    component named 'Text' → text_color. Every other body → body_color. This handles
    through-cut island fragments ('Sheet (1)' etc.) correctly without needing to
    enumerate body-name variants."""
    body_appr = _get_or_create_appearance(design, body_color)
    text_appr = _get_or_create_appearance(design, text_color)
    counts = {"body": 0, "text": 0}

    def walk(comp):
        is_text = comp.name == "Text"
        appr = text_appr if is_text else body_appr
        bucket = "text" if is_text else "body"
        if appr:
            for body in comp.bRepBodies:
                try:
                    body.appearance = appr
                    counts[bucket] += 1
                except Exception as ex:
                    _log(f"appearance assign {body.name} failed: {ex}")
        for occ in comp.occurrences:
            walk(occ.component)

    walk(parent_comp)
    _log(f"appearance applied: body={body_color}({counts['body']}) text={text_color}({counts['text']})")


def _offset_plane(comp, base_plane, offset, name):
    pl_in = comp.constructionPlanes.createInput()
    pl_in.setByOffset(base_plane, adsk.core.ValueInput.createByReal(offset))
    plane = comp.constructionPlanes.add(pl_in)
    plane.name = name
    plane.isLightBulbOn = False
    return plane


def _activate_component(target):
    """Walk the design's occurrence tree and activate the occurrence whose component == target."""
    design = adsk.fusion.Design.cast(_app.activeProduct)
    if not design:
        return False

    def walk(comp):
        for occ in comp.occurrences:
            if occ.component == target:
                occ.activate()
                return True
            if walk(occ.component):
                return True
        return False

    return walk(design.rootComponent)


def build_flat_sheet(comp, panels, depth, thickness):
    sk = comp.sketches.add(comp.xYConstructionPlane)
    sk.name = "flat_sheet_outline"
    total_len = sum(w for _, w in panels)
    sk.sketchCurves.sketchLines.addTwoPointRectangle(
        adsk.core.Point3D.create(0, 0, 0),
        adsk.core.Point3D.create(depth, total_len, 0)
    )

    prof = _largest_profile(sk)
    ext_in = comp.features.extrudeFeatures.createInput(
        prof, adsk.fusion.FeatureOperations.NewBodyFeatureOperation
    )
    ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(thickness))
    ext = comp.features.extrudeFeatures.add(ext_in)
    body = ext.bodies.item(0)
    body.name = "Sheet"
    return body


def _largest_profile(sketch):
    best = None
    best_a = -1
    for i in range(sketch.profiles.count):
        p = sketch.profiles.item(i)
        try:
            a = p.areaProperties().area
            if a > best_a:
                best_a = a
                best = p
        except Exception:
            continue
    return best


def add_hinge_grooves(comp, top_plane, panel_ranges, depth, t, p):
    hw = p["hinge_w"]
    ht = p["hinge_t"]
    groove_depth = t - ht
    if groove_depth <= 1e-5:
        return

    panel_list = list(panel_ranges.items())
    hinge_ys = [panel_list[i][1][1] for i in range(len(panel_list) - 1)]

    for i, hy in enumerate(hinge_ys):
        sk = comp.sketches.add(top_plane)
        sk.name = f"hinge_cut_{i}"
        y0 = hy - hw / 2
        y1 = hy + hw / 2
        sk.sketchCurves.sketchLines.addTwoPointRectangle(
            adsk.core.Point3D.create(0, y0, 0),
            adsk.core.Point3D.create(depth, y1, 0)
        )
        prof = _largest_profile(sk)
        if not prof:
            continue
        ext_in = comp.features.extrudeFeatures.createInput(
            prof, adsk.fusion.FeatureOperations.CutFeatureOperation
        )
        ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-groove_depth))
        comp.features.extrudeFeatures.add(ext_in)


def thin_sheet_ends(comp, top_plane, depth, total_len, t, thin_depth, recess_d):
    """Rabbet the outer face at both axial ends by `thin_depth` over a band of
    length `recess_d`. The end cap's outer ring (ring_w) fills this step and its
    clearance (clr) accounts for pocket fit — so thin_depth = ring_w + clr is
    computed by the caller. Assembled box has uniform outer width."""
    if thin_depth <= 1e-5 or recess_d <= 1e-5:
        return
    if thin_depth >= t:
        _log(f"thin_sheet_ends: thin_depth={thin_depth*10:.2f}mm >= t={t*10:.2f}mm; skipping")
        return
    if recess_d >= depth / 2:
        _log(f"thin_sheet_ends: recess_d={recess_d*10:.1f}mm >= depth/2; skipping")
        return

    bands = [
        (0.0, recess_d, "x0"),
        (depth - recess_d, depth, "x1"),
    ]
    for x0, x1, tag in bands:
        sk = comp.sketches.add(top_plane)
        sk.name = f"thin_band_{tag}"
        sk.sketchCurves.sketchLines.addTwoPointRectangle(
            adsk.core.Point3D.create(x0, 0, 0),
            adsk.core.Point3D.create(x1, total_len, 0),
        )
        prof = _largest_profile(sk)
        if not prof:
            continue
        ext_in = comp.features.extrudeFeatures.createInput(
            prof, adsk.fusion.FeatureOperations.CutFeatureOperation
        )
        ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-thin_depth))
        comp.features.extrudeFeatures.add(ext_in)
    _log(f"thin_sheet_ends: band={recess_d*10:.2f}mm depth={thin_depth*10:.2f}mm")


def _apply_text_style(text_input, p):
    """Set font name and bold/italic style on a SketchTextInput from params dict."""
    font = p.get("text_font", "Arial")
    if font:
        text_input.fontName = font
    style = 0
    if p.get("text_bold"):
        style |= adsk.fusion.TextStyles.TextStyleBold
    if p.get("text_italic"):
        style |= adsk.fusion.TextStyles.TextStyleItalic
    if style:
        text_input.textStyle = style


def _measure_text(comp, top_plane, text, th, p):
    """Create a throwaway sketch with the text at height `th`, return (actual_w, actual_h)."""
    sk = comp.sketches.add(top_plane)
    sk.name = "__text_measure__"
    sk.isVisible = False
    try:
        text_in = sk.sketchTexts.createInput2(text, th)
        _apply_text_style(text_in, p)
        text_in.setAsMultiLine(
            adsk.core.Point3D.create(0, 0, 0),
            adsk.core.Point3D.create(1000, th * 3, 0),
            adsk.core.HorizontalAlignments.LeftHorizontalAlignment,
            adsk.core.VerticalAlignments.BottomVerticalAlignment,
            0.0,
        )
        st = sk.sketchTexts.add(text_in)
        bb = st.boundingBox
        w = bb.maxPoint.x - bb.minPoint.x
        h = bb.maxPoint.y - bb.minPoint.y
        return w, h
    finally:
        try:
            sk.deleteMe()
        except Exception:
            pass


def _fit_text_height(comp, top_plane, text, requested_th, avail_x, avail_y, min_th, p):
    """Return the largest text height ≤ requested_th that keeps the rendered text inside
    (avail_x, avail_y). Measures iteratively because different fonts / glyph spacing can
    shift the effective aspect ratio."""
    SAFETY = 0.94
    target_w = avail_x * SAFETY
    target_h = avail_y * SAFETY

    th = requested_th
    for _ in range(4):
        w, h = _measure_text(comp, top_plane, text, th, p)
        if w <= 0 or h <= 0:
            return th
        if w <= target_w and h <= target_h:
            return th
        shrink = min(target_w / w, target_h / h)
        new_th = th * shrink * 0.99
        if new_th < min_th:
            _log(f"auto-fit: requested {requested_th*10:.1f}mm would shrink below 1mm; clamping")
            return min_th
        if abs(new_th - th) < 1e-4:
            break
        th = new_th

    if abs(th - requested_th) > 1e-4:
        _log(f"auto-fit: text_h {requested_th*10:.1f}mm -> {th*10:.1f}mm "
             f"(avail x={avail_x*10:.1f}mm y={avail_y*10:.1f}mm)")
    return th


def _autosize_depth(comp, plane, p):
    """Measure front + back text at requested height, return the minimum depth that
    comfortably fits whichever is longer (text + 8% margin each side)."""
    front = p["text_str"]
    back = (p.get("text_str_back") or "").strip() or front
    th = p["text_h"]
    if not front or th <= 0:
        return p["depth"]
    needed = 0.0
    for text in {front, back}:
        w, _ = _measure_text(comp, plane, text, th, p)
        if w > 0:
            needed = max(needed, w / 0.84)
    if needed <= 0:
        return p["depth"]
    _log(f"autosize depth: longer of {{'{front}','{back}'}} at {th*10:.1f}mm -> depth={needed*10:.1f}mm")
    return max(needed, 0.5)  # floor at 5mm


def add_text_bodies(comp, top_plane, panel_ranges, depth, t, p):
    """Create front/back text bodies. Three modes via p['text_mode']:

    - 'Flush Recess' (default): cut recess of depth `text_extrude` into the sheet top
      face, extrude a filler body flush with the top. Bodies kept distinct for
      multi-material slicing. `text_extrude < t` required.
    - 'Through-cut': cut is full sheet thickness; filler sits through the whole sheet
      (ideal for backlit translucent fills). `text_extrude` is ignored. Letters with
      closed inner regions (A, O, D, P, B, R, Q) leave disconnected sheet islands
      which are kept as separate sheet-colored bodies via `_normalize_sheet_names`
      — on the first print layer they fuse back into the main sheet.
    - 'Emboss': no cut; text body stands above the sheet top face, extruded +z by
      `text_extrude`. `text_extrude` can exceed `t` (raised signage look).

    Both sides share the same text height — the longer string drives auto-fit, the
    shorter centers at the same size so they match visually."""
    front_text = p["text_str"]
    back_text = (p.get("text_str_back") or "").strip() or front_text
    th = p["text_h"]
    mode = p.get("text_mode", "Flush Recess")
    user_te = p.get("text_extrude", 0.06)

    # Per-mode geometry. Body is extruded from top_plane; `body_up` goes +z (emboss),
    # `body_down` goes -z (fills the cut in the sheet).
    #   Flush Recess:          partial cut down, filler fills the recess.
    #   Through-cut:           full cut, filler fills through sheet.
    #   Emboss:                no cut, filler stands above sheet.
    #   Through-cut + Emboss:  full cut AND filler extends above sheet.
    if mode == "Flush Recess":
        cut_depth, body_down, body_up = user_te, user_te, 0.0
    elif mode == "Through-cut":
        cut_depth, body_down, body_up = t, t, 0.0
    elif mode == "Emboss":
        cut_depth, body_down, body_up = 0.0, 0.0, user_te
    elif mode == "Through-cut + Emboss":
        cut_depth, body_down, body_up = t, t, user_te
    else:
        _log(f"unknown text_mode={mode}; skipping")
        return

    if "front" not in panel_ranges or "back" not in panel_ranges:
        return
    if body_down <= 1e-5 and body_up <= 1e-5:
        _log(f"text body extent zero in mode={mode}; skipping")
        return
    if mode == "Flush Recess" and user_te >= t:
        _log(f"text_extrude={user_te} must be < sheet_t={t} in Flush Recess mode; skipping")
        return

    fy0, fy1, fcy = panel_ranges["front"]
    by0, by1, bcy = panel_ranges["back"]
    front_len = fy1 - fy0
    back_len = by1 - by0

    margin_x = depth * 0.08
    x0 = margin_x
    x1 = depth - margin_x
    avail_x = x1 - x0

    # Auto-fit by measuring the rendered text's bounding box. Fit both strings against
    # the tighter of the two panel lengths and take the smaller result — whichever text
    # is longer drives the final size, the shorter one inherits it and centers.
    MIN_TH = 0.1  # 1 mm floor
    if p.get("text_autofit", True):
        avail_y = min(front_len, back_len) * 0.76
        th_f = _fit_text_height(comp, top_plane, front_text, th, avail_x, avail_y, MIN_TH, p)
        th_b = th_f if back_text == front_text \
            else _fit_text_height(comp, top_plane, back_text, th, avail_x, avail_y, MIN_TH, p)
        th = min(th_f, th_b)

    def _panel_y_range(y_lo, y_hi):
        plen = y_hi - y_lo
        margin_y = min(plen * 0.12, th * 0.4)
        return y_lo + margin_y, y_hi - margin_y

    fy0_t, fy1_t = _panel_y_range(fy0, fy1)
    by0_t, by1_t = _panel_y_range(by0, by1)
    if fy1_t - fy0_t < th * 0.8 or by1_t - by0_t < th * 0.8:
        _log("front/back panel too short for text; skipping")
        return

    def _make_text_sketch(name, text_str, y_lo, y_hi):
        sk = comp.sketches.add(top_plane)
        sk.name = name
        text_in = sk.sketchTexts.createInput2(text_str, th)
        _apply_text_style(text_in, p)
        text_in.setAsMultiLine(
            adsk.core.Point3D.create(x0, y_lo, 0),
            adsk.core.Point3D.create(x1, y_hi, 0),
            adsk.core.HorizontalAlignments.CenterHorizontalAlignment,
            adsk.core.VerticalAlignments.MiddleVerticalAlignment,
            0.0,
        )
        st = sk.sketchTexts.add(text_in)
        coll = adsk.core.ObjectCollection.create()
        coll.add(st)
        return coll

    def _set_body_extent(extrude_input):
        """Single extrude that covers both sides of top_plane when needed."""
        vi = adsk.core.ValueInput.createByReal
        if body_up > 1e-5 and body_down > 1e-5:
            side_up = adsk.fusion.DistanceExtentDefinition.create(vi(body_up))
            side_dn = adsk.fusion.DistanceExtentDefinition.create(vi(body_down))
            extrude_input.setTwoSidesExtent(side_up, side_dn)
        elif body_up > 1e-5:
            extrude_input.setDistanceExtent(False, vi(+body_up))
        else:
            extrude_input.setDistanceExtent(False, vi(-body_down))

    # --- Front ---
    if cut_depth > 1e-5:
        cut_coll = _make_text_sketch("text_front_cut", front_text, fy0_t, fy1_t)
        cut_in = comp.features.extrudeFeatures.createInput(
            cut_coll, adsk.fusion.FeatureOperations.CutFeatureOperation
        )
        cut_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-cut_depth))
        try:
            comp.features.extrudeFeatures.add(cut_in)
        except Exception as e:
            _log(f"front text cut failed: {e}")
            return

    body_coll = _make_text_sketch("text_front_body", front_text, fy0_t, fy1_t)
    body_in = comp.features.extrudeFeatures.createInput(
        body_coll, adsk.fusion.FeatureOperations.NewBodyFeatureOperation
    )
    _set_body_extent(body_in)
    try:
        body_feat = comp.features.extrudeFeatures.add(body_in)
    except Exception as e:
        _log(f"front text body failed: {e}")
        return
    for i in range(body_feat.bodies.count):
        body_feat.bodies.item(i).name = f"Text_front_{i + 1}"

    # --- Back: extrude on back panel then rotate 180° about Z through back center so
    # the glyphs read correctly on the opposite face of the folded box. ---
    back_body_coll = _make_text_sketch("text_back_body", back_text, by0_t, by1_t)
    back_body_in = comp.features.extrudeFeatures.createInput(
        back_body_coll, adsk.fusion.FeatureOperations.NewBodyFeatureOperation
    )
    _set_body_extent(back_body_in)
    try:
        back_body_feat = comp.features.extrudeFeatures.add(back_body_in)
    except Exception as e:
        _log(f"back text body failed: {e}")
        return

    back_text_bodies = adsk.core.ObjectCollection.create()
    for i in range(back_body_feat.bodies.count):
        b = back_body_feat.bodies.item(i)
        b.name = f"Text_back_{i + 1}"
        back_text_bodies.add(b)

    rot = adsk.core.Matrix3D.create()
    rot.setToRotation(
        math.pi,
        adsk.core.Vector3D.create(0, 0, 1),
        adsk.core.Point3D.create(depth / 2.0, bcy, 0),
    )
    move_in = comp.features.moveFeatures.createInput2(back_text_bodies)
    move_in.defineAsFreeMove(rot)
    try:
        comp.features.moveFeatures.add(move_in)
    except Exception as e:
        _log(f"back text rotate failed: {e}")
        return

    # Cut the sheet using the rotated back text bodies (only when we actually cut).
    if cut_depth > 1e-5:
        sheet_body = _main_sheet_body(comp)
        if not sheet_body:
            _log("sheet body not found for back text cut")
            return
        combine_in = comp.features.combineFeatures.createInput(sheet_body, back_text_bodies)
        combine_in.operation = adsk.fusion.FeatureOperations.CutFeatureOperation
        combine_in.isKeepToolBodies = True
        try:
            comp.features.combineFeatures.add(combine_in)
        except Exception as e:
            _log(f"back text combine cut failed: {e}")

    # Full-depth cut splits off inner islands (inside of A/O/D/P/B/R/Q). Keep them
    # as separate sheet-colored bodies; just normalize the main sheet's name.
    if cut_depth >= t - 1e-5:
        _normalize_sheet_names(comp)

    # Move all text bodies into a "Text" sub-component so user can set material per group.
    text_occ = comp.occurrences.addNewComponent(adsk.core.Matrix3D.create())
    text_occ.component.name = "Text"

    all_text = [b for b in comp.bRepBodies if b.name.startswith("Text_")]
    for b in all_text:
        try:
            b.moveToComponent(text_occ)
        except Exception as e:
            _log(f"move {b.name} to Text component failed: {e}")


def _main_sheet_body(comp):
    """Return the largest body named 'Sheet' or 'Sheet (N)' in comp (the real sheet)."""
    candidates = [b for b in comp.bRepBodies if b.name == "Sheet" or b.name.startswith("Sheet (")]
    if not candidates:
        return None
    def _vol(b):
        try:
            return b.physicalProperties.volume
        except Exception:
            return 0.0
    return max(candidates, key=_vol)


def _normalize_sheet_names(comp):
    """After through-cut text, letters with closed inner regions (A, O, D, P, B, R, Q)
    leave disconnected material as separate 'Sheet (N)' bodies. We KEEP the islands
    (they sit flush inside each letter hole and fuse to the main sheet on the first
    print layer — visually the O's inner disk shows in sheet color).

    This only ensures the LARGEST body is named "Sheet" so anything keying off that
    name still works. The smaller bodies keep their auto-generated "Sheet (N)" names;
    appearance assignment keys off parent component, not body name, so they still
    get the sheet color."""
    sheets = [b for b in comp.bRepBodies if b.name == "Sheet" or b.name.startswith("Sheet (")]
    if not sheets:
        return
    if len(sheets) == 1:
        if sheets[0].name != "Sheet":
            sheets[0].name = "Sheet"
        return

    def _vol(b):
        try:
            return b.physicalProperties.volume
        except Exception:
            return 0.0
    sheets.sort(key=_vol, reverse=True)

    # Find the largest. If it's not already called "Sheet", swap names with whichever
    # body currently holds that name (Fusion rejects duplicates, so a direct rename
    # to "Sheet" would silently no-op when another body already owns it).
    main = sheets[0]
    if main.name == "Sheet":
        return
    current_sheet = next((b for b in sheets if b.name == "Sheet"), None)
    main_old_name = main.name
    if current_sheet:
        current_sheet.name = f"__tmp_{main_old_name}"
        main.name = "Sheet"
        # Re-assign a stable auto-style name to the demoted one
        current_sheet.name = main_old_name
    else:
        main.name = "Sheet"
    _log(f"through-cut: kept {len(sheets)} sheet body/bodies (main + islands)")


def _fillet_cap_corners(cap_comp, plate_body, cap_t, recess_d, corner_r, ring_w):
    """Round the cap's vertical corner edges to match the folded sheet's fold arcs.
    - Plate outer corners (full-height edges z=0..cap_t): fillet with corner_r
    - Pocket outer corners (partial-height edges in recess): fillet with max(corner_r - ring_w, 0.2)
      to keep the ring roughly constant radial thickness around corners
    - Pocket inner corners: left sharp (inner fold arc ≈ 0 for our fold axis model)"""
    tol = 1e-4
    plate_vert = []
    pocket_vert = []
    for e in plate_body.edges:
        sv = e.startVertex.geometry
        ev = e.endVertex.geometry
        if abs(sv.x - ev.x) > tol or abs(sv.y - ev.y) > tol:
            continue
        z_min, z_max = min(sv.z, ev.z), max(sv.z, ev.z)
        if abs(z_min) < tol and abs(z_max - cap_t) < tol:
            plate_vert.append(e)
        elif abs(z_min - (cap_t - recess_d)) < tol and abs(z_max - cap_t) < tol:
            pocket_vert.append(e)

    # Classify pocket verts into outer vs inner by radial distance from body centroid
    bb = plate_body.boundingBox
    cx = (bb.minPoint.x + bb.maxPoint.x) / 2
    cy = (bb.minPoint.y + bb.maxPoint.y) / 2
    pocket_outer_edges = []
    if pocket_vert:
        dists = [(e, math.hypot(e.startVertex.geometry.x - cx,
                                e.startVertex.geometry.y - cy)) for e in pocket_vert]
        # Split at the median — outer half get filleted
        dists.sort(key=lambda x: -x[1])
        half = len(dists) // 2
        pocket_outer_edges = [x[0] for x in dists[:half]]

    def _apply(edges, r):
        if not edges or r <= 1e-4:
            return
        coll = adsk.core.ObjectCollection.create()
        for ed in edges:
            coll.add(ed)
        fi = cap_comp.features.filletFeatures.createInput()
        fi.addConstantRadiusEdgeSet(coll, adsk.core.ValueInput.createByReal(r), True)
        try:
            cap_comp.features.filletFeatures.add(fi)
        except Exception as ex:
            _log(f"cap fillet r={r*10:.2f}mm failed: {ex}")

    _apply(plate_vert, corner_r)
    _apply(pocket_outer_edges, max(corner_r - ring_w, 0.02))
    _log(f"cap fillet: plate={corner_r*10:.2f}mm pocket_outer={max(corner_r-ring_w, 0.02)*10:.2f}mm "
         f"(n_plate={len(plate_vert)}, n_pocket_outer={len(pocket_outer_edges)})")


def build_end_caps(parent, panels, depth, t, p):
    """Flush-fit end caps. Outer outline matches the folded box's outer envelope, so
    cap + body share the same width. Pocket is an ANNULAR groove that receives the
    sheet's thinned axial end (see thin_sheet_ends): the ring between the outer and
    inner loops carries the outer rim material + inner plug for stability.

    Geometry (fold axis assumed on sheet inner face):
      plate_outer     = profile + t              (flush with box outer)
      pocket_outer    = profile + (t - ring_w)   (ring material thickness = ring_w)
      pocket_inner    = profile - clr            (inner plug, slight clearance)
      sheet thin_depth = ring_w + clr            (set by thin_sheet_ends caller)
    ring_w is user-set so the outer rim is always printable (default 0.8mm = 2 perimeters).
    """
    cap_t = p["endcap_t"]
    recess_d = p["endcap_recess"]
    clr = p["endcap_clr"]
    ring_w = p["endcap_ring_w"]
    profile = p["profile"]
    top_w = p["top_w"]
    bot_w = p["bot_w"]
    height = p["height"]

    if recess_d >= cap_t:
        recess_d = cap_t - 0.1  # keep at least 1 mm back wall

    cap_occ = parent.occurrences.addNewComponent(adsk.core.Matrix3D.create())
    cap_comp = cap_occ.component
    cap_comp.name = "EndCaps"

    base_outline = profile_outline(profile, top_w, bot_w, height)
    plate_outline = outset_outline(base_outline, t)
    pocket_outer = outset_outline(base_outline, t - ring_w)
    pocket_inner = inset_outline(base_outline, clr)

    if not plate_outline or not pocket_outer or not pocket_inner:
        _log("end cap: outline offset failed (profile too small for clearances)")
        return
    _log(f"endcap: plate=+{t*10:.2f}mm ring_w={ring_w*10:.2f}mm "
         f"pocket_outer=+{(t-ring_w)*10:.2f}mm pocket_inner=-{clr*10:.2f}mm "
         f"recess={recess_d*10:.2f}mm (thin_depth={(ring_w+clr)*10:.2f}mm)")

    gap = 0.5
    max_py = max(pt[1] for pt in plate_outline)
    min_py = min(pt[1] for pt in plate_outline)
    plate_h = max_py - min_py

    def _draw_loop(sk_lines, outline, y_off):
        n = len(outline)
        for i in range(n):
            ax, az = outline[i]
            bx, bz = outline[(i + 1) % n]
            sk_lines.addByTwoPoints(
                adsk.core.Point3D.create(ax, az + y_off, 0),
                adsk.core.Point3D.create(bx, bz + y_off, 0),
            )

    for idx in range(2):
        y_off = -(max_py + gap + idx * (plate_h + gap))

        # Plate: solid extrude of the outer outline
        sk = cap_comp.sketches.add(cap_comp.xYConstructionPlane)
        sk.name = f"endcap_{idx}_plate"
        _draw_loop(sk.sketchCurves.sketchLines, plate_outline, y_off)

        plate_prof = _largest_profile(sk)
        if not plate_prof:
            _log(f"endcap {idx}: no plate profile")
            continue
        ext_in = cap_comp.features.extrudeFeatures.createInput(
            plate_prof, adsk.fusion.FeatureOperations.NewBodyFeatureOperation
        )
        ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(cap_t))
        plate_ext = cap_comp.features.extrudeFeatures.add(ext_in)
        plate_body = plate_ext.bodies.item(0)
        plate_body.name = f"EndCap_{idx + 1}"

        # Annular pocket: two nested closed loops. Fusion creates two profiles from
        # this — the inner disk (inside pocket_inner) and the ring (between the loops).
        # Pick the ring by finding a profile with profileLoops.count == 2.
        sk2 = cap_comp.sketches.add(cap_comp.xYConstructionPlane)
        sk2.name = f"endcap_{idx}_pocket"
        _draw_loop(sk2.sketchCurves.sketchLines, pocket_outer, y_off)
        _draw_loop(sk2.sketchCurves.sketchLines, pocket_inner, y_off)

        ring_prof = None
        for i in range(sk2.profiles.count):
            pr = sk2.profiles.item(i)
            if pr.profileLoops.count == 2:
                ring_prof = pr
                break
        if not ring_prof:
            _log(f"endcap {idx}: annular pocket profile not found")
            continue

        cut_in = cap_comp.features.extrudeFeatures.createInput(
            ring_prof, adsk.fusion.FeatureOperations.CutFeatureOperation
        )
        cut_in.startExtent = adsk.fusion.OffsetStartDefinition.create(
            adsk.core.ValueInput.createByReal(cap_t - recess_d)
        )
        cut_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(recess_d))
        try:
            cap_comp.features.extrudeFeatures.add(cut_in)
        except Exception as e:
            _log(f"endcap {idx} pocket cut failed: {e}")

        corner_r = p.get("endcap_corner_r", 0.0)
        if corner_r > 1e-4:
            _fillet_cap_corners(cap_comp, plate_body, cap_t, recess_d, corner_r, ring_w)


def profile_outline(profile, top_w, bot_w, height):
    """Profile outline points (x, z) going CCW; x is horizontal (width), z is vertical (height)."""
    if profile == "Triangle":
        return [
            (-bot_w / 2, 0),
            (+bot_w / 2, 0),
            (0, height),
        ]
    if profile in ("Square", "Trapezoid"):
        return [
            (-bot_w / 2, 0),
            (+bot_w / 2, 0),
            (+top_w / 2, height),
            (-top_w / 2, height),
        ]
    raise ValueError(profile)


def _line_intersection(p1, p2, p3, p4):
    """Intersection of infinite lines (p1->p2) and (p3->p4). Returns None if parallel."""
    x1, y1 = p1; x2, y2 = p2
    x3, y3 = p3; x4, y4 = p4
    den = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    if abs(den) < 1e-10:
        return None
    tnum = (x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)
    t = tnum / den
    return (x1 + t * (x2 - x1), y1 + t * (y2 - y1))


def _polygon_offset(outline, d):
    """True perpendicular polygon offset by d (positive = outward, negative = inward).
    Assumes CCW vertex order (as returned by profile_outline). Each edge is shifted by
    its outward normal * d, and adjacent shifted edges are intersected to get new
    vertices. Returns None if the offset collapses the polygon or any edge degenerates."""
    if not outline or len(outline) < 3:
        return None
    n = len(outline)
    shifted = []
    for i in range(n):
        p0, p1 = outline[i], outline[(i + 1) % n]
        dx, dy = p1[0] - p0[0], p1[1] - p0[1]
        length = math.hypot(dx, dy)
        if length < 1e-9:
            return None
        # CCW polygon: outward normal is (dy, -dx) / length
        nx, ny = dy / length, -dx / length
        sp0 = (p0[0] + nx * d, p0[1] + ny * d)
        sp1 = (p1[0] + nx * d, p1[1] + ny * d)
        shifted.append((sp0, sp1))

    verts = []
    for i in range(n):
        a, b = shifted[(i - 1) % n]
        c, e = shifted[i]
        ix = _line_intersection(a, b, c, e)
        if ix is None:
            return None
        verts.append(ix)

    # Collapse guard (inset only): each new vertex must be at inward distance >= |d|
    # from EVERY original edge, not just its two adjacent ones. Inset past the
    # inradius produces a topologically valid CCW polygon whose vertices are on the
    # wrong side of some edge (or past the opposite edge). Convex outset cannot
    # collapse so we skip the check there.
    if d < 0:
        need = -d - 1e-6
        for v in verts:
            for i in range(n):
                p0 = outline[i]
                p1 = outline[(i + 1) % n]
                edx, edy = p1[0] - p0[0], p1[1] - p0[1]
                length = math.hypot(edx, edy)
                if length < 1e-9:
                    continue
                nx, ny = -edy / length, edx / length  # CCW inward normal
                dist = (v[0] - p0[0]) * nx + (v[1] - p0[1]) * ny
                if dist < need:
                    return None
        # Degenerate at exactly the inradius: distance check passes (dist == |d|) but
        # the polygon collapses to ~zero area. Reject anything under 0.01 mm².
        area2 = 0.0
        for i in range(n):
            x0, y0 = verts[i]
            x1, y1 = verts[(i + 1) % n]
            area2 += x0 * y1 - x1 * y0
        if area2 <= 1e-4:
            return None
    return verts


def inset_outline(outline, inset):
    """Inward perpendicular polygon offset (CCW input). Returns None if inset would
    collapse the polygon (inset >= inradius)."""
    if not outline:
        return None
    if inset <= 0:
        return list(outline)
    return _polygon_offset(outline, -inset)


def outset_outline(outline, offset):
    """Outward perpendicular polygon offset (CCW input)."""
    if not outline:
        return None
    if offset <= 0:
        return list(outline)
    return _polygon_offset(outline, offset)
