# The app's one loading screen. Appended to PyInstaller's own splash script by
# AnimatedSplash in BoxOfScraps.spec, which fills in the @PLACEHOLDERS@ at build
# time. PyInstaller's part has already created the window and .root.canvas, and
# set up the connection File.py talks to.
#
# This file must stay pure ASCII: AnimatedSplash in the spec explains why and refuses
# to build otherwise. Write anything else as a Tcl backslash-u escape.
#
# This runs in the bootloader process, on its own thread, so it keeps animating
# no matter how busy the app's Python process is.
#
# Progress comes from two real sources, both arriving through status_text:
#   - while unpacking, the bootloader writes the name of each file it extracts;
#   - after that, File.py sends "@floor ceiling" as it reaches each stage.
# Between those events the display drifts toward the ceiling without reaching
# it, so a slow stage slows the bar down instead of pretending to be finished.

# Share of the bar that unpacking fills, and how far it may drift while Python
# starts up afterwards. File.py's _splash_stage calls continue from these numbers.
set anim_extract_share @EXTRACT_SHARE@
set anim_boot_ceiling  @BOOT_CEILING@

# Each file's cost in units; files missing from this table cost 1.
array set anim_cost {@FILE_COSTS@}
set anim_cost_total @COST_TOTAL@
set anim_cost_done  0
set anim_extracting 1

set anim_shown   0.0
set anim_floor   0.0
set anim_ceiling 0.0

image create photo anim_owl_img  -data {@OWL_PNG@}
image create photo anim_logo_img -data {@LOGO_PNG@}

# The logo runs right to left with the owl chasing it. The pass is tied to the
# load percentage, so it crosses exactly once and leaves the left edge at 100%.
set anim_ground 92
set anim_gap    100
set anim_logo_start [expr {$image_width + [image width anim_logo_img] / 2.0 - 12}]
set anim_logo_end   [expr {-[image width anim_owl_img] / 2.0 - 10 - $anim_gap}]

.root.canvas create image -200 $anim_ground -image anim_logo_img -tags anim_logo
.root.canvas create image -200 $anim_ground -image anim_owl_img  -tags anim_owl
.root.canvas create text [expr {$image_width / 2}] 170 \
    -text "Loading\u2026 0%" -fill "#073b5e" -font {{Segoe UI} 11} -tags anim_pct

proc anim_on_status {args} {
    global status_text anim_cost anim_cost_total anim_cost_done anim_extracting
    global anim_extract_share anim_boot_ceiling anim_floor anim_ceiling

    # A stage reported by File.py. Unpacking is over once Python is talking.
    if {[string index $status_text 0] eq "@"} {
        set anim_extracting 0
        lassign [string range $status_text 1 end] floor ceiling
        set anim_floor   [expr {max($anim_floor, $floor)}]
        set anim_ceiling [expr {max($anim_floor, $ceiling)}]
        return
    }
    if {!$anim_extracting} {
        return
    }

    # A file the bootloader is unpacking.
    if {[info exists anim_cost($status_text)]} {
        incr anim_cost_done $anim_cost($status_text)
    } else {
        incr anim_cost_done
    }
    set share [expr {min(1.0, double($anim_cost_done) / $anim_cost_total)}]
    set anim_floor   [expr {$share * $anim_extract_share}]
    set anim_ceiling $anim_floor
    if {$share >= 0.995} {
        # Nothing reports progress until Python connects, so drift through that.
        set anim_extracting 0
        set anim_floor   $anim_extract_share
        set anim_ceiling $anim_boot_ceiling
    }
}
trace add variable status_text write anim_on_status

proc anim_tick {} {
    global anim_shown anim_floor anim_ceiling anim_ground anim_gap
    global anim_logo_start anim_logo_end

    # Glide quickly up to what has been proven and creep toward the ceiling, taking
    # whichever is the bigger step. Both ease in without ever arriving, so neither
    # can be made to wait for the other to finish. 100 is the last thing shown before
    # the window appears, so there is nothing to glide to.
    if {$anim_floor >= 100} {
        set anim_shown 100.0
    } else {
        set anim_shown [expr {$anim_shown + max(0.0,
                              ($anim_floor - $anim_shown) * 0.15,
                              ($anim_ceiling - $anim_shown) * 0.08)}]
    }

    set logo_x [expr {$anim_logo_start
                      + ($anim_logo_end - $anim_logo_start) * $anim_shown / 100.0}]
    # A slight hop as they run, the owl out of phase with the logo. This is the
    # part that keeps moving even while the percentage is waiting on something.
    set phase [expr {[clock milliseconds] * 0.019}]
    .root.canvas coords anim_logo $logo_x \
        [expr {$anim_ground - 8 * abs(sin($phase))}]
    .root.canvas coords anim_owl [expr {$logo_x + $anim_gap}] \
        [expr {$anim_ground - 8 * abs(sin($phase + 0.9))}]
    .root.canvas itemconfigure anim_pct \
        -text "Loading\u2026 [expr {int($anim_shown)}]%"

    after 20 anim_tick
}
anim_tick
