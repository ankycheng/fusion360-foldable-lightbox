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

# Custom event used to defer STEP import outside the command execute context.
# Fusion documents importToTarget as unsupported inside command events — calling
# it there can crash Fusion. A custom event fires on the main thread AFTER the
# command loop unwinds, so importToTarget runs safely there.
GARMIN_IMPORT_EVENT_ID = "FoldableLightbox_doGarminImport"
_garmin_import_event = None  # keep a module-level handle for stop() cleanup
# Queue of pending Garmin STEP imports. Each entry is a dict with:
#   - step_path   : absolute path to the Garmin connector STEP file
#   - depth_cm    : sheet depth (for X centering)
#   - base_y_mid_cm : base panel Y midpoint (for Y centering)
#   - sheet_t_cm  : sheet thickness (unused for Z now — we want Z=0)
#   - pre_names   : set() of root.occurrences names before the import, so the
#                   handler can identify the newly-added occurrence
_pending_garmin_imports = []

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
    global _app, _ui, _garmin_import_event
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

        # Register the deferred-STEP-import custom event. It's fired from
        # add_garmin_mount() inside onExecute, but Fusion queues the dispatch
        # so the handler runs outside command scope where importToTarget works.
        # registerCustomEvent returns None if the id was already registered
        # (e.g. add-in was Stop-Run'd without a full Fusion restart) — in that
        # case, unregister first then re-register so we rebind to a fresh handler.
        ev = _app.registerCustomEvent(GARMIN_IMPORT_EVENT_ID)
        if ev is None:
            _app.unregisterCustomEvent(GARMIN_IMPORT_EVENT_ID)
            ev = _app.registerCustomEvent(GARMIN_IMPORT_EVENT_ID)
        if ev is not None:
            garmin_handler = GarminImportHandler()
            added = ev.add(garmin_handler)
            _handlers.append(garmin_handler)
            _garmin_import_event = ev
            _log(f"run(): garmin custom event registered, handler.add returned {added}")
        else:
            _log("run(): registerCustomEvent returned None twice — deferred Garmin import will not work")

        ws = _ui.workspaces.itemById("FusionSolidEnvironment")
        panel = ws.toolbarPanels.itemById("SolidCreatePanel")
        if not panel.controls.itemById(CMD_ID):
            panel.controls.addCommand(cmd_def)
    except Exception:
        if _ui:
            _ui.messageBox(f"Foldable Lightbox failed to start:\n{traceback.format_exc()}")


def stop(context):
    global _handlers, _garmin_import_event
    try:
        ws = _ui.workspaces.itemById("FusionSolidEnvironment")
        panel = ws.toolbarPanels.itemById("SolidCreatePanel")
        ctrl = panel.controls.itemById(CMD_ID)
        if ctrl:
            ctrl.deleteMe()
        cmd_def = _ui.commandDefinitions.itemById(CMD_ID)
        if cmd_def:
            cmd_def.deleteMe()
        if _app is not None:
            try:
                _app.unregisterCustomEvent(GARMIN_IMPORT_EVENT_ID)
            except Exception:
                pass
    except Exception:
        pass
    _handlers = []
    _garmin_import_event = None


_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "foldable_lightbox.log")


def _log(msg):
    line = f"[FoldableLightbox] {msg}"
    try:
        palette = _ui.palettes.itemById("TextCommands")
        if palette:
            palette.writeText(line)
    except Exception:
        pass
    try:
        import datetime as _dt
        with open(_LOG_PATH, "a") as f:
            f.write(f"{_dt.datetime.now().isoformat(timespec='seconds')}  {line}\n")
    except Exception:
        pass


class GarminImportHandler(adsk.core.CustomEventHandler):
    """Deferred STEP import handler. Fires *after* the Foldable Lightbox
    command has fully finished, so importManager.importToTarget runs outside
    command scope (where it is supported). Payload is a JSON dict with:
      - mount_token: entityToken of the GarminMount placeholder occurrence
      - step_path:   absolute path to the Garmin connector STEP file
    """
    def notify(self, args):
        _log(f"garmin deferred: handler.notify() ENTERED; queue_size={len(_pending_garmin_imports)}")
        if not _pending_garmin_imports:
            _log("garmin deferred: queue empty, nothing to import")
            return
        while _pending_garmin_imports:
            params = _pending_garmin_imports.pop(0)
            _log(f"garmin deferred: processing params={list(params.keys())}")
            self._run_one(params)

    def _run_one(self, params):
        try:
            step_path = params["step_path"]
            depth = params["depth_cm"]
            base_y_mid = params["base_y_mid_cm"]
            sheet_t = params.get("sheet_t_cm", 0.0)
            pre_names = set(params.get("pre_names", []))

            if not os.path.exists(step_path):
                _log(f"garmin deferred: STEP file missing at {step_path}")
                return

            app = adsk.core.Application.get()
            design = app.activeProduct
            if not design or design.classType() != adsk.fusion.Design.classType():
                _log(f"garmin deferred: active product is not a Design")
                return
            root = design.rootComponent

            # Perform the STEP import at root. STEP will create a new linked
            # child document and place its reference occurrence at root; the
            # `target` argument is effectively ignored for STEP files.
            im = app.importManager
            opts = im.createSTEPImportOptions(step_path)
            opts.isViewFit = False
            im.importToTarget(opts, root)

            # Diff root.occurrences against the pre-import snapshot to find
            # the one we just added.
            new_occ = None
            for i in range(root.occurrences.count):
                occ = root.occurrences.item(i)
                if occ.name not in pre_names:
                    new_occ = occ
                    break
            if new_occ is None:
                _log("garmin deferred: could not identify new imported occurrence (names all pre-existing)")
                return
            _log(f"garmin deferred: new occurrence '{new_occ.name}' detected")

            # The imported body sits in the CHILD component in its source STEP
            # frame (Fusion does NOT auto-rotate Y-up → Z-up for this STEP).
            # Verified via face-normal query:
            #   X = disc width (ear-to-ear, ~28.8mm)
            #   Y = disc THICKNESS axis; largest planar face has normal +Y
            #       (the flange *mounting face* with the 3 screw holes — this
            #        is the face that must rest on the sheet top)
            #   Z = disc depth (perpendicular to ears, ~24.6mm)
            #   -Y side carries the bayonet twist-lock tabs ("接口") which
            #        must point AWAY from the sheet (into +Z).
            # Rotation −90° around X: (x, y, z) → (x, z, −y). This maps:
            #   +Y face (mounting, srcMaxY) → post-rot minZ = −srcMaxY  (BOTTOM)
            #   −Y face (tabs,     srcMinY) → post-rot maxZ = −srcMinY  (TOP)
            # Then translate so the mounting face lands on Z = sheet_t (sheet
            # top surface) with the tabs protruding to Z = sheet_t + thickness.
            body = None
            if new_occ.component.bRepBodies.count > 0:
                body = new_occ.component.bRepBodies.item(0)
            if body is None:
                _log("garmin deferred: imported occurrence has no body")
                return
            try:
                body.name = "GarminConnector"
            except Exception:
                pass

            bb = body.boundingBox
            src_cx = (bb.minPoint.x + bb.maxPoint.x) / 2.0
            src_cz = (bb.minPoint.z + bb.maxPoint.z) / 2.0
            src_max_y = bb.maxPoint.y

            Tx = depth / 2.0 - src_cx
            Ty = base_y_mid - src_cz  # post-rot Y = +src_Z, so subtract src_cz
            Tz = sheet_t + src_max_y  # mounting face (post-rot minZ = −srcMaxY) lands on Z = sheet_t

            _log(f"garmin deferred: body local bbox X({bb.minPoint.x*10:.1f}..{bb.maxPoint.x*10:.1f}) "
                 f"Y({bb.minPoint.y*10:.1f}..{bb.maxPoint.y*10:.1f}) "
                 f"Z({bb.minPoint.z*10:.1f}..{bb.maxPoint.z*10:.1f})mm; "
                 f"rot−90°X, translation=({Tx*10:.1f},{Ty*10:.1f},{Tz*10:.1f})mm")

            # Find the Lightbox_{profile} parent occurrence at root level by name.
            # build_lightbox creates exactly one per run, so name-prefix scan is
            # unambiguous. root.occurrences.item() returns a natively-typed
            # Occurrence, unlike findEntityByToken which returns a Base wrapper
            # that SWIG rejects in moveToComponent.
            parent_occ = None
            _log(f"garmin deferred: scanning {root.occurrences.count} root occurrences for Lightbox_*")
            try:
                for i in range(root.occurrences.count):
                    occ = root.occurrences.item(i)
                    _log(f"garmin deferred:   [{i}] name='{occ.name}'")
                    if occ.name.startswith("Lightbox_"):
                        parent_occ = occ
                        break
            except Exception as e:
                _log(f"garmin deferred: root.occurrences scan failed: {e}")

            # CRITICAL ORDERING: moveToComponent returns a NEW Occurrence proxy
            # in the target's assembly context. The old new_occ reference becomes
            # stale (its assemblyContext still reports root, and transform set on
            # it gets dropped when Fusion rebuilds the occurrence under the new
            # parent). So:
            #   1. Call moveToComponent FIRST
            #   2. Switch to using the returned occurrence
            #   3. THEN set transform on the new reference
            # Setting transform before move caused the Garmin to land at its raw
            # STEP bbox position (sheet's upper-right corner, ~(180, 126)mm).
            target_occ = new_occ
            if parent_occ is not None:
                try:
                    result = new_occ.moveToComponent(parent_occ)
                    if isinstance(result, adsk.fusion.Occurrence):
                        target_occ = result
                        _log(f"garmin deferred: moved into '{parent_occ.name}', using returned occurrence for transform")
                    else:
                        _log(f"garmin deferred: moveToComponent returned unexpected type "
                             f"({type(result).__name__}); using original reference")
                except Exception as e:
                    _log(f"garmin deferred: moveToComponent failed: {e}")
            else:
                _log("garmin deferred: no Lightbox_* occurrence at root; leaving connector at root")

            # Apply rotation + translation. Parent (Lightbox_*) has identity
            # transform, so these values are interpreted identically in parent
            # local and world coords — no math change needed.
            mat = adsk.core.Matrix3D.create()
            mat.setToRotation(-math.pi / 2.0,
                              adsk.core.Vector3D.create(1, 0, 0),
                              adsk.core.Point3D.create(0, 0, 0))
            mat.translation = adsk.core.Vector3D.create(Tx, Ty, Tz)
            try:
                target_occ.transform = mat
                _log(f"garmin deferred: applied transform to '{target_occ.name}'; "
                     f"center ({(Tx+src_cx)*10:.1f},{(Ty+src_cz)*10:.1f},"
                     f"{(Tz-src_max_y)*10:.1f}..{(Tz-bb.minPoint.y)*10:.1f})mm")
            except Exception as e:
                _log(f"garmin deferred: setting transform on target_occ failed: {e}")

            try:
                if design.snapshots.hasPendingSnapshot:
                    design.snapshots.add()
            except Exception:
                pass
        except Exception:
            _log(f"garmin deferred handler failed:\n{traceback.format_exc()}")


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

            g_switch = inputs.addGroupCommandInput("g_switch", "Switch Boss + Hole (one end cap)")
            g_switch.children.addBoolValueInput("switch_hole", "Add Switch Mounting Boss", True, "",
                s.get("switch_hole", False))
            g_switch.children.addValueInput("switch_boss_d", "Boss Diameter", "mm",
                _vi(s, "switch_boss_d", "15 mm"))
            g_switch.children.addValueInput("switch_boss_h", "Boss Height (inward)", "mm",
                _vi(s, "switch_boss_h", "5 mm"))
            g_switch.children.addValueInput("switch_hole_d", "Through-hole Diameter (clearance)", "mm",
                _vi(s, "switch_hole_d", "6.5 mm"))
            g_switch.children.addBoolValueInput("switch_tap_thread",
                "Tap 1/4-40 UNS-2B Internal Thread (overrides Ø, uses 5.7mm tap drill)", True, "",
                s.get("switch_tap_thread", False))

            g_cutouts = inputs.addGroupCommandInput("g_cutouts", "End Cap Cutouts (plain hole / USB-C)")
            g_cutouts.children.addBoolValueInput("cap1_plain_hole",
                "Cap 1: Plain Through-Hole (alternative to switch boss)", True, "",
                s.get("cap1_plain_hole", False))
            g_cutouts.children.addValueInput("cap1_plain_hole_d", "Plain Hole Diameter", "mm",
                _vi(s, "cap1_plain_hole_d", "8.1 mm"))
            g_cutouts.children.addBoolValueInput("cap2_usbc_port",
                "Cap 2: USB-C Port Cutout", True, "",
                s.get("cap2_usbc_port", False))
            g_cutouts.children.addValueInput("cap2_usbc_w", "USB-C Cutout Width", "mm",
                _vi(s, "cap2_usbc_w", "8.9 mm"))
            g_cutouts.children.addValueInput("cap2_usbc_h", "USB-C Cutout Height", "mm",
                _vi(s, "cap2_usbc_h", "3.3 mm"))
            g_cutouts.children.addValueInput("cap2_usbc_r", "USB-C Corner Radius", "mm",
                _vi(s, "cap2_usbc_r", "1.65 mm"))
            g_cutouts.children.addValueInput("cap2_usbc_y_off", "USB-C Vertical Offset (−down, +up)", "mm",
                _vi(s, "cap2_usbc_y_off", "0 mm"))
            g_cutouts.children.addBoolValueInput("cap2_pcb_slot",
                "Cap 2: PCB Slot (blind pocket on inner face, below USB-C)", True, "",
                s.get("cap2_pcb_slot", False))
            g_cutouts.children.addValueInput("cap2_pcb_slot_w", "PCB Slot Width", "mm",
                _vi(s, "cap2_pcb_slot_w", "17.5 mm"))
            g_cutouts.children.addValueInput("cap2_pcb_slot_h", "PCB Slot Height (PCB thickness)", "mm",
                _vi(s, "cap2_pcb_slot_h", "1.2 mm"))
            g_cutouts.children.addValueInput("cap2_pcb_slot_d", "PCB Slot Depth (into cap)", "mm",
                _vi(s, "cap2_pcb_slot_d", "2 mm"))
            g_cutouts.children.addValueInput("cap2_pcb_slot_gap", "Gap below USB-C (slot top to USB-C bottom)", "mm",
                _vi(s, "cap2_pcb_slot_gap", "0 mm"))

            g_mount = inputs.addGroupCommandInput("g_mount", "Bottom Mount Holes (cross pattern)")
            g_mount.children.addBoolValueInput("mount_holes", "Add 4-hole Mount Pattern", True, "",
                s.get("mount_holes", False))
            g_mount.children.addValueInput("mount_hole_d", "Hole Diameter (clearance)", "mm",
                _vi(s, "mount_hole_d", "3.0 mm"))
            g_mount.children.addValueInput("mount_spacing", "Cross Spacing (opposite-hole distance)", "mm",
                _vi(s, "mount_spacing", "18 mm"))
            mount_style_dd = g_mount.children.addDropDownCommandInput(
                "mount_style", "Fixation Style",
                adsk.core.DropDownStyles.TextListDropDownStyle,
            )
            _MOUNT_STYLES = [
                "Clearance only",
                "Tap M3 Thread",
                "Hex Nut Pocket",
                "Garmin Quarter-Turn Connector",
            ]
            sel_mount_style = s.get("mount_style", "Clearance only")
            if sel_mount_style not in _MOUNT_STYLES:
                sel_mount_style = "Clearance only"
            for ms in _MOUNT_STYLES:
                mount_style_dd.listItems.add(ms, ms == sel_mount_style)
            g_mount.children.addValueInput("mount_pad_t", "Inner Pad Thickness (for thread / nut)", "mm",
                _vi(s, "mount_pad_t", "4 mm"))

            g_tab = inputs.addGroupCommandInput("g_tab", "Bottom Seam Tab")
            g_tab.children.addBoolValueInput("seam_tab", "Add Seam Tab (folds up against back panel)", True, "",
                s.get("seam_tab", False))
            g_tab.children.addValueInput("seam_tab_h", "Tab Height", "mm",
                _vi(s, "seam_tab_h", "4 mm"))
            g_tab.children.addValueInput("seam_tab_t", "Tab Thickness", "mm",
                _vi(s, "seam_tab_t", "0.3 mm"))

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
                "switch_hole":     _item(inp, "switch_hole").value,
                "switch_boss_d":   _item(inp, "switch_boss_d").value,
                "switch_boss_h":   _item(inp, "switch_boss_h").value,
                "switch_hole_d":   _item(inp, "switch_hole_d").value,
                "switch_tap_thread": _item(inp, "switch_tap_thread").value,
                "cap1_plain_hole":   _item(inp, "cap1_plain_hole").value,
                "cap1_plain_hole_d": _item(inp, "cap1_plain_hole_d").value,
                "cap2_usbc_port":    _item(inp, "cap2_usbc_port").value,
                "cap2_usbc_w":       _item(inp, "cap2_usbc_w").value,
                "cap2_usbc_h":       _item(inp, "cap2_usbc_h").value,
                "cap2_usbc_r":       _item(inp, "cap2_usbc_r").value,
                "cap2_usbc_y_off":   _item(inp, "cap2_usbc_y_off").value,
                "cap2_pcb_slot":     _item(inp, "cap2_pcb_slot").value,
                "cap2_pcb_slot_w":   _item(inp, "cap2_pcb_slot_w").value,
                "cap2_pcb_slot_h":   _item(inp, "cap2_pcb_slot_h").value,
                "cap2_pcb_slot_d":   _item(inp, "cap2_pcb_slot_d").value,
                "cap2_pcb_slot_gap": _item(inp, "cap2_pcb_slot_gap").value,
                "mount_holes":     _item(inp, "mount_holes").value,
                "mount_hole_d":    _item(inp, "mount_hole_d").value,
                "mount_spacing":   _item(inp, "mount_spacing").value,
                "mount_style":     _item(inp, "mount_style").selectedItem.name,
                "mount_pad_t":     _item(inp, "mount_pad_t").value,
                "seam_tab":        _item(inp, "seam_tab").value,
                "seam_tab_h":      _item(inp, "seam_tab_h").value,
                "seam_tab_t":      _item(inp, "seam_tab_t").value,
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

    # This add-in needs to create multiple sub-components (FlatSheet, EndCaps,
    # GarminMount, …). A Part Design document only allows ONE component, so
    # detect that case up-front and give the user a clear fix instead of a
    # raw Fusion exception.
    try:
        parent_occ = root.occurrences.addNewComponent(adsk.core.Matrix3D.create())
    except RuntimeError as e:
        if "Part Design" in str(e) or "only contain one component" in str(e):
            msg = (
                "Foldable Lightbox needs an Assembly document (supports multiple components).\n"
                "The active document is a Part Design, which can hold only one component.\n\n"
                "Fix: File → New Design from File → New Design (Assembly), or use the\n"
                "Design workspace's 'New Assembly' command, then re-run this add-in."
            )
            _log("build_lightbox: aborted — Part Design document detected")
            try:
                if _ui:
                    _ui.messageBox(msg, "Foldable Lightbox — Wrong Document Type")
            except Exception:
                pass
            return
        raise
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
    tab_h = p["seam_tab_h"] if p.get("seam_tab") else 0.0
    tab_t = p.get("seam_tab_t", 0.0) if p.get("seam_tab") else 0.0

    _log(f"profile={p['profile']} panels={[(n, round(w, 3)) for n, w in panels]} depth={depth} tab_h={tab_h} tab_t={tab_t}")

    sheet_body = build_flat_sheet(sheet_comp, panels, depth, t, tab_h)

    top_plane = _offset_plane(sheet_comp, sheet_comp.xYConstructionPlane, t, "top_sketch_plane")
    if tab_h > 0 and tab_t > 0 and tab_t < t - 1e-5:
        thin_seam_tab(sheet_comp, top_plane, depth, tab_h, t, tab_t)
    add_hinge_grooves(sheet_comp, top_plane, panel_ranges, depth, t, p, tab_h, tab_t)

    if p["text_str"] and p["text_h"] > 0:
        add_text_bodies(sheet_comp, top_plane, panel_ranges, depth, t, p)

    if p.get("mount_holes"):
        if p.get("mount_style") == "Garmin Quarter-Turn Connector":
            # Garmin mode: no drilled holes / pad / nuts — the connector body
            # is the entire fixation and sits flush on the base panel bottom.
            # Pass parent_occ (not parent) so we can grab a root-rooted proxy
            # of the new GarminMount occurrence — required for entityToken.
            add_garmin_mount(parent_occ, panel_ranges, depth, t)
        else:
            add_mount_holes(sheet_comp, top_plane, panel_ranges, depth, t, p)

    if p["endcaps"]:
        total_len = sum(w for _, w in panels)
        thin_depth = p["endcap_ring_w"] + p["endcap_clr"]
        thin_sheet_ends(sheet_comp, top_plane, depth, total_len, t,
                        thin_depth, p["endcap_recess"], tab_h)
        build_end_caps(parent, panels, depth, t, p, tab_h)

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


def build_flat_sheet(comp, panels, depth, thickness, tab_h=0.0):
    sk = comp.sketches.add(comp.xYConstructionPlane)
    sk.name = "flat_sheet_outline"
    total_len = sum(w for _, w in panels)
    sk.sketchCurves.sketchLines.addTwoPointRectangle(
        adsk.core.Point3D.create(0, -tab_h, 0),
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


def thin_seam_tab(comp, top_plane, depth, tab_h, t, tab_t):
    """Rabbet the tab region (Y < 0) from the top face so the tab is only
    `tab_t` thick instead of full sheet thickness `t`. The step at Y=0 acts
    as a living hinge between tab and base."""
    thin_by = t - tab_t
    if thin_by <= 1e-5 or tab_h <= 0:
        return
    sk = comp.sketches.add(top_plane)
    sk.name = "tab_thin"
    sk.sketchCurves.sketchLines.addTwoPointRectangle(
        adsk.core.Point3D.create(-0.01, -tab_h - 0.01, 0),
        adsk.core.Point3D.create(depth + 0.01, 0, 0),
    )
    prof = _largest_profile(sk)
    if not prof:
        return
    ext_in = comp.features.extrudeFeatures.createInput(
        prof, adsk.fusion.FeatureOperations.CutFeatureOperation
    )
    ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-thin_by))
    comp.features.extrudeFeatures.add(ext_in)


def add_hinge_grooves(comp, top_plane, panel_ranges, depth, t, p, tab_h=0.0, tab_t=0.0):
    hw = p["hinge_w"]
    ht = p["hinge_t"]
    groove_depth = t - ht
    if groove_depth <= 1e-5:
        return

    panel_list = list(panel_ranges.items())
    hinge_ys = [panel_list[i][1][1] for i in range(len(panel_list) - 1)]
    # If the seam tab is already thinned to ≤ sheet thickness it acts as its
    # own living hinge at Y=0, so no explicit groove is needed there.
    tab_full_t = tab_h > 0 and (tab_t <= 0 or tab_t >= t - 1e-5)
    if tab_full_t:
        hinge_ys = [0.0] + hinge_ys

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


def thin_sheet_ends(comp, top_plane, depth, total_len, t, thin_depth, recess_d, tab_h=0.0):
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
            adsk.core.Point3D.create(x0, -tab_h, 0),
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


def add_mount_holes(comp, top_plane, panel_ranges, depth, t, p):
    """Drill clearance holes through 'base' panel in a cross pattern.
    Optionally add an inner pad + tapped thread or hex nut pocket."""
    base = panel_ranges.get("base")
    if not base:
        _log("mount holes: no base panel; skipping")
        return
    y_min, y_max, y_mid = base
    x_mid = depth / 2.0
    user_r = p["mount_hole_d"] / 2.0
    half = p["mount_spacing"] / 2.0
    style = p.get("mount_style", "Clearance only")
    pad_t = p.get("mount_pad_t", 0.4) if style != "Clearance only" else 0.0

    panel_half_y = (y_max - y_min) / 2.0
    x_fits = (half + user_r) <= depth / 2.0
    y_fits = (half + user_r) <= panel_half_y

    if not x_fits and not y_fits:
        min_needed_mm = (p["mount_spacing"] + p["mount_hole_d"]) * 10
        msg = (
            f"Mount holes skipped: base panel ({(y_max-y_min)*10:.1f}×{depth*10:.1f}mm) "
            f"too small for spacing {p['mount_spacing']*10:.1f}mm + Ø{p['mount_hole_d']*10:.1f}mm clearance. "
            f"Need at least one side ≥ {min_needed_mm:.1f}mm. Increase Bottom Width / Depth or reduce spacing."
        )
        _log(msg)
        try:
            if _ui:
                _ui.messageBox(msg, "Foldable Lightbox — Mount Holes")
        except Exception:
            pass
        return

    centers = []
    if x_fits:
        centers += [(x_mid + half, y_mid), (x_mid - half, y_mid)]
    if y_fits:
        centers += [(x_mid, y_mid + half), (x_mid, y_mid - half)]

    sheet_body = comp.bRepBodies.itemByName("Sheet")

    # ── Build inner pad (Join to Sheet) ──
    if pad_t > 0 and sheet_body:
        margin = 0.2  # cm
        xs = [c[0] for c in centers]
        ys = [c[1] for c in centers]
        px0 = max(0, min(xs) - user_r - margin)
        px1 = min(depth, max(xs) + user_r + margin)
        py0 = max(y_min, min(ys) - user_r - margin)
        py1 = min(y_max, max(ys) + user_r + margin)
        sk_pad = comp.sketches.add(top_plane)
        sk_pad.name = "mount_pad"
        sk_pad.sketchCurves.sketchLines.addTwoPointRectangle(
            adsk.core.Point3D.create(px0, py0, 0),
            adsk.core.Point3D.create(px1, py1, 0),
        )
        prof_pad = _largest_profile(sk_pad)
        if prof_pad:
            ext_in = comp.features.extrudeFeatures.createInput(
                prof_pad, adsk.fusion.FeatureOperations.JoinFeatureOperation)
            ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(pad_t))
            try:
                ext_in.participantBodies = [sheet_body]
            except Exception:
                coll = adsk.core.ObjectCollection.create()
                coll.add(sheet_body)
                ext_in.participantBodies = coll
            comp.features.extrudeFeatures.add(ext_in)

    # ── Drill clearance / tap-drill holes through sheet + pad ──
    # For M3 thread, hole ø must be the tap drill (2.5mm for M3x0.5), not user clearance.
    cut_r = 0.125 if style == "Tap M3 Thread" else user_r
    if pad_t > 0:
        drill_plane = _offset_plane(comp, comp.xYConstructionPlane, t + pad_t,
                                    "mount_drill_plane")
    else:
        drill_plane = top_plane
    total_depth = t + pad_t

    sk_h = comp.sketches.add(drill_plane)
    sk_h.name = "mount_holes"
    for cx, cy in centers:
        sk_h.sketchCurves.sketchCircles.addByCenterRadius(
            adsk.core.Point3D.create(cx, cy, 0), cut_r)

    cut_feats = []
    for i in range(sk_h.profiles.count):
        prof = sk_h.profiles.item(i)
        ext_in = comp.features.extrudeFeatures.createInput(
            prof, adsk.fusion.FeatureOperations.CutFeatureOperation
        )
        ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-total_depth))
        cut_feats.append(comp.features.extrudeFeatures.add(ext_in))

    # ── Apply M3×0.5 internal thread on each hole ──
    if style == "Tap M3 Thread":
        tf = comp.features.threadFeatures
        ti = tf.createThreadInfo(True, "ISO Metric profile", "M3x0.5", "6H")
        for cf in cut_feats:
            cyl = next((f for f in cf.sideFaces if isinstance(f.geometry, adsk.core.Cylinder)), None)
            if not cyl:
                continue
            face_coll = adsk.core.ObjectCollection.create()
            face_coll.add(cyl)
            t_in = tf.createInput(face_coll, ti)
            t_in.isModeled = True
            try:
                tf.add(t_in)
            except Exception as e:
                _log(f"mount hole M3 thread fail: {e}")

    # ── Cut hex nut pockets from sheet BOTTOM (Z=0) upward into the pad.
    # Opening on sheet bottom face → after folding this faces box INSIDE, so nut
    # drops in from inside the box; screw threads in from outside (through pad).
    # Trade-off: hex ceiling is an overhang when printing flat → needs support.
    if style == "Hex Nut Pocket":
        aflat = 0.55    # M3 DIN 934 across-flats 5.5mm
        nut_d = 0.26    # pocket depth 2.6mm (nut 2.4mm + 0.2mm clearance)
        r_hex = aflat / (2 * math.cos(math.radians(30)))
        sk_hex = comp.sketches.add(comp.xYConstructionPlane)
        sk_hex.name = "mount_hex_pockets"
        for cx, cy in centers:
            pts = [
                (cx + r_hex * math.cos(math.radians(30 + 60*i)),
                 cy + r_hex * math.sin(math.radians(30 + 60*i)))
                for i in range(6)
            ]
            for i in range(6):
                sk_hex.sketchCurves.sketchLines.addByTwoPoints(
                    adsk.core.Point3D.create(pts[i][0], pts[i][1], 0),
                    adsk.core.Point3D.create(pts[(i+1) % 6][0], pts[(i+1) % 6][1], 0),
                )
        for i in range(sk_hex.profiles.count):
            prof = sk_hex.profiles.item(i)
            ext_in = comp.features.extrudeFeatures.createInput(
                prof, adsk.fusion.FeatureOperations.CutFeatureOperation)
            ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(nut_d))
            comp.features.extrudeFeatures.add(ext_in)

    n = len(centers)
    axis = "cross" if n == 4 else ("depth" if x_fits else "width")
    _log(f"mount holes: {n}x {style}, spacing={p['mount_spacing']*10:.1f}mm, pad={pad_t*10:.1f}mm, axis={axis}")


def add_garmin_mount(parent_occ, panel_ranges, depth, t):
    """Queue a deferred Garmin connector STEP import + placement.

    Fusion's importManager.importToTarget must NOT be called inside a command
    event (it can crash Fusion). Observed actual behaviour: for a STEP file
    it always creates a *new root-level occurrence* linking to an auto-
    imported child document — regardless of the `target` component passed
    in — and auto-rotates the body from STEP Y-up to Fusion Z-up.

    So here we just record the intent (including sheet geometry needed to
    centre the body on the base panel) and fire a custom event. The paired
    GarminImportHandler runs AFTER onExecute returns, imports the STEP at
    root, finds the newly-created occurrence, and sets its transform to
    land the body on the base panel's bottom face (Z=0) centred at
    (depth/2, base_y_mid).
    """
    base = panel_ranges.get("base")
    if not base:
        _log("garmin mount: no base panel; skipping")
        return
    _, _, base_y_mid = base

    step_path = os.path.join(os.path.dirname(__file__), "garmin_connector_male.step")
    if not os.path.exists(step_path):
        msg = (f"Garmin connector STEP file missing at\n  {step_path}\n\n"
               "Skipping Garmin mount. Ship the STEP file alongside the add-in "
               "to enable this feature.")
        _log(f"garmin mount: STEP missing at {step_path}; skipping")
        try:
            if _ui:
                _ui.messageBox(msg, "Foldable Lightbox — Garmin Mount")
        except Exception:
            pass
        return

    # Snapshot current root occurrences — the handler will diff against this
    # to detect the newly-imported occurrence.
    try:
        design = adsk.core.Application.get().activeProduct
        root = design.rootComponent
        pre_names = {root.occurrences.item(i).name for i in range(root.occurrences.count)}
    except Exception as e:
        _log(f"garmin mount: failed to snapshot root occurrences: {e}")
        pre_names = set()

    # Capture the Lightbox parent occurrence's entityToken so the deferred
    # handler can resolve it post-import and nest the Garmin connector inside
    # the Lightbox group instead of leaving it at root.
    parent_token = ""
    try:
        parent_token = parent_occ.entityToken or ""
    except Exception as e:
        _log(f"garmin mount: failed to capture parent entityToken: {e}")

    params = {
        "step_path": step_path,
        "depth_cm": depth,
        "base_y_mid_cm": base_y_mid,
        "sheet_t_cm": t,
        "pre_names": list(pre_names),
        "parent_token": parent_token,
    }
    _pending_garmin_imports.append(params)

    try:
        app = adsk.core.Application.get()
        fired = app.fireCustomEvent(GARMIN_IMPORT_EVENT_ID, "")
        _log(f"garmin mount: enqueued import depth={depth*10:.1f}mm "
             f"base_y_mid={base_y_mid*10:.1f}mm t={t*10:.2f}mm, "
             f"pre_occ_count={len(pre_names)}, fireCustomEvent={fired}")
    except Exception as e:
        _log(f"garmin mount: failed to fire event: {e}\n{traceback.format_exc()}")


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
    shift the effective aspect ratio.
    SAFETY = 0.97 leaves a 3% buffer for measurement-vs-layout variance; tighter than
    0.94 so text fills more of the panel when autosize-depth is on."""
    SAFETY = 0.97
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
    comfortably fits whichever is longer. The divisor combines the full margin chain
    that add_text_bodies + _fit_text_height will apply downstream so autofit doesn't
    shrink the text further once the depth is sized:
      avail_x = depth * (1 - 2*MARGIN_X_FRAC)   where MARGIN_X_FRAC = 0.04
      target_w = avail_x * SAFETY               where SAFETY = 0.97 in _fit_text_height
    → depth_needed = w / (0.92 * 0.97) ≈ w / 0.8924
    """
    front = p["text_str"]
    back = (p.get("text_str_back") or "").strip() or front
    th = p["text_h"]
    if not front or th <= 0:
        return p["depth"]
    needed = 0.0
    usable_frac = 0.92 * 0.97  # keep in sync with add_text_bodies margin + _fit_text_height SAFETY
    for text in {front, back}:
        w, _ = _measure_text(comp, plane, text, th, p)
        if w > 0:
            needed = max(needed, w / usable_frac)
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

    # Margin fractions kept tight so text fills most of the panel (user feedback:
    # "上下左右還有空間"). Keep in sync with _autosize_depth's usable_frac.
    margin_x = depth * 0.04
    x0 = margin_x
    x1 = depth - margin_x
    avail_x = x1 - x0

    # Auto-fit by measuring the rendered text's bounding box. Fit both strings against
    # the tighter of the two panel lengths and take the smaller result — whichever text
    # is longer drives the final size, the shorter one inherits it and centers.
    MIN_TH = 0.1  # 1 mm floor
    if p.get("text_autofit", True):
        avail_y = min(front_len, back_len) * 0.88
        th_f = _fit_text_height(comp, top_plane, front_text, th, avail_x, avail_y, MIN_TH, p)
        th_b = th_f if back_text == front_text \
            else _fit_text_height(comp, top_plane, back_text, th, avail_x, avail_y, MIN_TH, p)
        th = min(th_f, th_b)

    def _panel_y_range(y_lo, y_hi):
        plen = y_hi - y_lo
        margin_y = min(plen * 0.06, th * 0.25)
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
    all_text_bodies = []
    for i in range(body_feat.bodies.count):
        b = body_feat.bodies.item(i)
        b.name = f"Text_front_{i + 1}"
        all_text_bodies.append(b)

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
        all_text_bodies.append(b)

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
    # Use the direct refs we captured at extrude time — comp.bRepBodies sometimes
    # returns proxies that moveToComponent rejects, leaving bodies in the parent.
    text_occ = comp.occurrences.addNewComponent(adsk.core.Matrix3D.create())
    text_occ.component.name = "Text"

    moved = 0
    for b in all_text_bodies:
        try:
            b.moveToComponent(text_occ)
            moved += 1
        except Exception as e:
            _log(f"move {b.name} to Text component failed: {e}")
    # Fallback: any body still named Text_* in sheet_comp (e.g. islands) goes too.
    for b in list(comp.bRepBodies):
        if b.name.startswith("Text_"):
            try:
                b.moveToComponent(text_occ)
                moved += 1
            except Exception as e:
                _log(f"move fallback {b.name} failed: {e}")
    _log(f"text move: {moved}/{len(all_text_bodies)} direct + fallback moved into Text component")


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
    leave disconnected material as separate 'Sheet (N)' bodies. These islands sit
    inside each letter hole and fuse to the main sheet on the first print layer —
    visually the letter's inner area shows in sheet color.

    Caveat: STL export emits the islands as separate connected components.
    Bambu Studio / PrusaSlicer / Cura auto-split STL by connectivity and will
    show each island as its own floating object. Workarounds for the user:
      (1) Export as 3MF instead of STL (preserves per-body grouping in the
          slicer), or
      (2) After STL import, use the slicer's "Merge parts" / "Unite" action to
          combine the islands with the main sheet — they'll still print as
          intended, with the first layer fusing everything.

    Boolean-joining disjoint bodies inside Fusion is not directly supported
    by CombineFeature (Fusion errors "Some input argument is invalid") and
    synthesising a thin bridge is fragile in the parametric timeline, so we
    leave the islands as separate bodies and only fix up the naming: largest
    body -> "Sheet" so downstream name-keyed code keeps working."""
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
    main = sheets[0]
    if main.name != "Sheet":
        current_sheet = next((b for b in sheets if b.name == "Sheet"), None)
        main_old_name = main.name
        if current_sheet is not None and current_sheet is not main:
            current_sheet.name = f"__tmp_{main_old_name}"
            main.name = "Sheet"
            current_sheet.name = main_old_name
        else:
            main.name = "Sheet"
    _log(f"through-cut: {len(sheets)} sheet body/bodies (main + {len(sheets)-1} island"
         f"{'s' if len(sheets)-1 != 1 else ''}); islands stay as separate bodies — "
         f"export to 3MF or 'Merge parts' in slicer to treat as one object")


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


def _add_switch_boss_hole(comp, cap_body, y_off, outline, cap_t,
                          boss_d, boss_h, hole_d, tap_thread=False):
    """Add cylindrical boss on cap's top (inner) face at profile bbox center,
    then drill a through-hole (boss + cap). outline is list of (x, z) tuples —
    z is used as sketch-Y here (matches _draw_loop convention).
    If tap_thread is True, hole_d is overridden with 5.3mm tap drill and an
    M6×0.75 internal thread feature is added on the hole face."""
    if tap_thread:
        # 1/4-40 UNS-2B internal thread (datasheet: 100SP3T1B1M2QE toggle switch
        # uses 1/4-40 UNS-2A bushing). Tap drill ≈ major − pitch = 6.35 − 0.635
        # = 5.715mm; use 5.7mm for a forgiving self-tap fit in FDM plastic.
        hole_d = 0.57
    xs = [pt[0] for pt in outline]
    zs = [pt[1] for pt in outline]
    cx = (min(xs) + max(xs)) / 2.0
    cy = (min(zs) + max(zs)) / 2.0 + y_off

    top_plane = _offset_plane(comp, comp.xYConstructionPlane, cap_t,
                              f"switch_plane_{id(cap_body)}")

    sk_b = comp.sketches.add(top_plane)
    sk_b.name = "switch_boss_circle"
    sk_b.sketchCurves.sketchCircles.addByCenterRadius(
        adsk.core.Point3D.create(cx, cy, 0), boss_d / 2.0)
    prof_b = _largest_profile(sk_b)
    if not prof_b:
        _log("switch boss: no profile")
        return
    ext_in = comp.features.extrudeFeatures.createInput(
        prof_b, adsk.fusion.FeatureOperations.JoinFeatureOperation
    )
    ext_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(boss_h))
    try:
        ext_in.participantBodies = [cap_body]
    except Exception:
        coll = adsk.core.ObjectCollection.create()
        coll.add(cap_body)
        ext_in.participantBodies = coll
    comp.features.extrudeFeatures.add(ext_in)

    # Hole sketch must sit on the boss's *top* face so a -Z cut passes through
    # boss + cap in one feature. Using top_plane (Z=cap_t) would miss the boss
    # because the boss grows in +Z above that plane.
    hole_plane = _offset_plane(comp, comp.xYConstructionPlane, cap_t + boss_h,
                               f"switch_hole_plane_{id(cap_body)}")
    sk_h = comp.sketches.add(hole_plane)
    sk_h.name = "switch_hole_circle"
    sk_h.sketchCurves.sketchCircles.addByCenterRadius(
        adsk.core.Point3D.create(cx, cy, 0), hole_d / 2.0)
    prof_h = _largest_profile(sk_h)
    if not prof_h:
        _log("switch hole: no profile")
        return
    cut_in = comp.features.extrudeFeatures.createInput(
        prof_h, adsk.fusion.FeatureOperations.CutFeatureOperation
    )
    cut_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-(boss_h + cap_t)))
    cut_feat = comp.features.extrudeFeatures.add(cut_in)

    if tap_thread:
        hole_face = None
        for f in cut_feat.sideFaces:
            if isinstance(f.geometry, adsk.core.Cylinder):
                hole_face = f
                break
        if hole_face:
            # Mini toggle switch bushings (e.g. C&K 100SP3T1B1M2QE) use
            # 1/4-40 UNS-2A external threads — the mating internal thread we
            # model here is 1/4-40 UNS-2B. Fusion's thread library lists this
            # under "ANSI Unified Screw Threads"; fall back to "M6x0.75" if the
            # library designation is unavailable on the user's install.
            tf = comp.features.threadFeatures
            attempts = [
                ("ANSI Unified Screw Threads", "1/4-40 UNS", "2B"),
                ("ANSI Unified Screw Threads", "1/4-40", "2B"),
                ("ISO Metric profile", "M6x0.75", "6H"),  # last-resort fallback
            ]
            thread_added = False
            for thread_type, designation, cls in attempts:
                try:
                    ti = tf.createThreadInfo(True, thread_type, designation, cls)
                    face_coll = adsk.core.ObjectCollection.create()
                    face_coll.add(hole_face)
                    t_in = tf.createInput(face_coll, ti)
                    t_in.isModeled = True
                    tf.add(t_in)
                    _log(f"switch hole: boss Ø{boss_d*10:.2f}×{boss_h*10:.2f}mm, "
                         f"{designation} ({thread_type} cls {cls}) tapped Ø{hole_d*10:.2f}mm "
                         f"through {cap_t*10:.2f}+{boss_h*10:.2f}mm")
                    thread_added = True
                    break
                except Exception as e:
                    _log(f"switch thread: {thread_type}/{designation}/{cls} failed: {e}")
            if not thread_added:
                _log(f"switch thread: all thread designations failed; leaving as plain Ø{hole_d*10:.2f}mm hole")
            return
        else:
            _log(f"switch thread: no cylindrical hole face found; leaving as plain Ø{hole_d*10:.2f}mm hole")

    _log(f"switch hole: boss Ø{boss_d*10:.2f}×{boss_h*10:.2f}mm, hole Ø{hole_d*10:.2f}mm through {cap_t*10:.2f}+{boss_h*10:.2f}mm")


def _add_cap_plain_hole(comp, cap_body, y_off, outline, cap_t, hole_d):
    """Simple Ø through-hole on cap plate, centered in profile bbox. No boss, no thread.
    Alternative to _add_switch_boss_hole when the user just wants a clean clearance hole
    (e.g. Ø8.1mm for a cable gland or bushing).
    Sketches on a plane ABOVE the cap top so the plate outline isn't inherited as a
    profile — cuts -Z by (cap_t + overshoot) to punch through the full plate."""
    xs = [pt[0] for pt in outline]
    zs = [pt[1] for pt in outline]
    cx = (min(xs) + max(xs)) / 2.0
    cy = (min(zs) + max(zs)) / 2.0 + y_off

    hole_plane = _offset_plane(comp, comp.xYConstructionPlane, cap_t + 0.1,
                               f"cap_plain_hole_plane_{id(cap_body)}")
    sk = comp.sketches.add(hole_plane)
    sk.name = f"cap_plain_hole_{id(cap_body)}"
    sk.sketchCurves.sketchCircles.addByCenterRadius(
        adsk.core.Point3D.create(cx, cy, 0), hole_d / 2.0)
    prof = _largest_profile(sk)
    if not prof:
        _log("cap plain hole: no profile")
        return
    cut_in = comp.features.extrudeFeatures.createInput(
        prof, adsk.fusion.FeatureOperations.CutFeatureOperation
    )
    cut_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-(cap_t + 0.2)))
    try:
        comp.features.extrudeFeatures.add(cut_in)
        _log(f"cap plain hole: Ø{hole_d*10:.2f}mm through {cap_t*10:.2f}mm plate")
    except Exception as e:
        _log(f"cap plain hole failed: {e}")


def _add_cap_usbc_cutout(comp, cap_body, y_off, outline, cap_t, w, h, r, y_shift=0.0):
    """Rounded-rectangle cutout (USB-C port) through cap plate, centered in profile
    bbox. Width w is along the profile horizontal axis (sketch X), height h is along
    the vertical axis (sketch Y). y_shift shifts the port along the profile vertical
    axis (negative = toward bottom). Corner radius r is clamped to min(w/2, h/2) − eps
    so straight edges keep positive length (at r = h/2 the shape becomes a stadium)."""
    xs = [pt[0] for pt in outline]
    zs = [pt[1] for pt in outline]
    cx = (min(xs) + max(xs)) / 2.0
    cy = (min(zs) + max(zs)) / 2.0 + y_off + y_shift

    eps = 1e-3  # 10 µm, keeps degenerate edges non-zero for Fusion
    r = max(0.0, min(r, w / 2.0 - eps, h / 2.0 - eps))

    hole_plane = _offset_plane(comp, comp.xYConstructionPlane, cap_t + 0.1,
                               f"cap_usbc_plane_{id(cap_body)}")
    sk = comp.sketches.add(hole_plane)
    sk.name = f"cap_usbc_cutout_{id(cap_body)}"

    hw = w / 2.0
    hh = h / 2.0
    lines = sk.sketchCurves.sketchLines
    arcs = sk.sketchCurves.sketchArcs

    if r < 1e-4:
        lines.addTwoPointRectangle(
            adsk.core.Point3D.create(cx - hw, cy - hh, 0),
            adsk.core.Point3D.create(cx + hw, cy + hh, 0),
        )
    else:
        # 4 straight edges (shortened by r at each end)
        lines.addByTwoPoints(
            adsk.core.Point3D.create(cx - hw + r, cy + hh, 0),
            adsk.core.Point3D.create(cx + hw - r, cy + hh, 0),
        )
        lines.addByTwoPoints(
            adsk.core.Point3D.create(cx - hw + r, cy - hh, 0),
            adsk.core.Point3D.create(cx + hw - r, cy - hh, 0),
        )
        lines.addByTwoPoints(
            adsk.core.Point3D.create(cx - hw, cy - hh + r, 0),
            adsk.core.Point3D.create(cx - hw, cy + hh - r, 0),
        )
        lines.addByTwoPoints(
            adsk.core.Point3D.create(cx + hw, cy - hh + r, 0),
            adsk.core.Point3D.create(cx + hw, cy + hh - r, 0),
        )
        # 4 corner arcs (+90° CCW sweep from each start)
        # TR: east → north
        arcs.addByCenterStartSweep(
            adsk.core.Point3D.create(cx + hw - r, cy + hh - r, 0),
            adsk.core.Point3D.create(cx + hw, cy + hh - r, 0),
            math.pi / 2.0,
        )
        # TL: north → west
        arcs.addByCenterStartSweep(
            adsk.core.Point3D.create(cx - hw + r, cy + hh - r, 0),
            adsk.core.Point3D.create(cx - hw + r, cy + hh, 0),
            math.pi / 2.0,
        )
        # BL: west → south
        arcs.addByCenterStartSweep(
            adsk.core.Point3D.create(cx - hw + r, cy - hh + r, 0),
            adsk.core.Point3D.create(cx - hw, cy - hh + r, 0),
            math.pi / 2.0,
        )
        # BR: south → east
        arcs.addByCenterStartSweep(
            adsk.core.Point3D.create(cx + hw - r, cy - hh + r, 0),
            adsk.core.Point3D.create(cx + hw - r, cy - hh, 0),
            math.pi / 2.0,
        )

    prof = _largest_profile(sk)
    if not prof:
        _log("cap USB-C: no profile")
        return
    cut_in = comp.features.extrudeFeatures.createInput(
        prof, adsk.fusion.FeatureOperations.CutFeatureOperation
    )
    cut_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-(cap_t + 0.2)))
    try:
        comp.features.extrudeFeatures.add(cut_in)
        _log(f"cap USB-C: {w*10:.2f}×{h*10:.2f}mm r={r*10:.2f}mm through {cap_t*10:.2f}mm plate")
    except Exception as e:
        _log(f"cap USB-C failed: {e}")


def _add_cap_pcb_slot(comp, cap_body, y_off, outline, cap_t,
                      usbc_h, w, h, depth, gap, usbc_y_shift=0.0):
    # Min back-wall kept under the slot so the plate stays closed on the outer face.
    min_back_wall = 0.08  # 0.8 mm
    xs = [pt[0] for pt in outline]
    zs = [pt[1] for pt in outline]
    cx = (min(xs) + max(xs)) / 2.0
    profile_cy = (min(zs) + max(zs)) / 2.0 + y_off + usbc_y_shift

    usbc_bottom_y = profile_cy - usbc_h / 2.0
    slot_top_y = usbc_bottom_y - gap
    slot_cy = slot_top_y - h / 2.0

    max_depth = cap_t - min_back_wall
    if depth > max_depth:
        _log(f"cap PCB slot: depth {depth*10:.2f}mm > cap_t-0.8mm "
             f"({max_depth*10:.2f}mm); trimmed")
        depth = max_depth
    if depth <= 0 or h <= 0 or w <= 0:
        _log("cap PCB slot: non-positive size, skipped")
        return

    slot_plane = _offset_plane(comp, comp.xYConstructionPlane, cap_t,
                               f"cap_pcb_slot_plane_{id(cap_body)}")
    sk = comp.sketches.add(slot_plane)
    sk.name = f"cap_pcb_slot_{id(cap_body)}"

    hw = w / 2.0
    hh = h / 2.0
    sk.sketchCurves.sketchLines.addTwoPointRectangle(
        adsk.core.Point3D.create(cx - hw, slot_cy - hh, 0),
        adsk.core.Point3D.create(cx + hw, slot_cy + hh, 0),
    )
    prof = _largest_profile(sk)
    if not prof:
        _log("cap PCB slot: no profile")
        return
    cut_in = comp.features.extrudeFeatures.createInput(
        prof, adsk.fusion.FeatureOperations.CutFeatureOperation
    )
    # Sketch plane is at z=cap_t (inner face); cut toward -Z into the plate.
    cut_in.setDistanceExtent(False, adsk.core.ValueInput.createByReal(-depth))
    try:
        comp.features.extrudeFeatures.add(cut_in)
        _log(f"cap PCB slot: {w*10:.2f}×{h*10:.2f}mm depth {depth*10:.2f}mm, "
             f"gap below USB-C {gap*10:.2f}mm")
    except Exception as e:
        _log(f"cap PCB slot failed: {e}")


def build_end_caps(parent, panels, depth, t, p, tab_h=0.0):
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

    gap = 1.0  # 10 mm between sheet (incl. tab) and first end cap, and between caps
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
        y_off = -(tab_h + max_py + gap + idx * (plate_h + gap))

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

        if idx == 0 and p.get("switch_hole"):
            _add_switch_boss_hole(cap_comp, plate_body, y_off, plate_outline,
                                  cap_t, p["switch_boss_d"], p["switch_boss_h"],
                                  p["switch_hole_d"],
                                  tap_thread=p.get("switch_tap_thread", False))
        if idx == 0 and p.get("cap1_plain_hole"):
            _add_cap_plain_hole(cap_comp, plate_body, y_off, plate_outline,
                                cap_t, p["cap1_plain_hole_d"])
        if idx == 1 and p.get("cap2_usbc_port"):
            usbc_y_shift = p.get("cap2_usbc_y_off", 0.0)
            _add_cap_usbc_cutout(cap_comp, plate_body, y_off, plate_outline,
                                 cap_t,
                                 p["cap2_usbc_w"], p["cap2_usbc_h"],
                                 p["cap2_usbc_r"],
                                 y_shift=usbc_y_shift)
            if p.get("cap2_pcb_slot"):
                _add_cap_pcb_slot(cap_comp, plate_body, y_off, plate_outline,
                                  cap_t,
                                  p["cap2_usbc_h"],
                                  p["cap2_pcb_slot_w"],
                                  p["cap2_pcb_slot_h"],
                                  p["cap2_pcb_slot_d"],
                                  p["cap2_pcb_slot_gap"],
                                  usbc_y_shift=usbc_y_shift)


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
