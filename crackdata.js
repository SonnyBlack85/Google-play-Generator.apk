-- Align selected cells across selection
-- Copyright under GPL by Mark Grimes
-- Saving with '\sca' in the filename creates Shortcut: Crtl+Shift+a

tell application "Microsoft Excel"
    --activate
    tell range (get address selection) of active sheet
        if (get count columns) > 1 or (get count rows) > 1 then
            if (get horizontal alignment) is horizontal align center across selection then
                set horizontal alignment to horizontal align general
            else
                set horizontal alignment to horizontal align center across selection
            end if
        else
            if (get horizontal alignment) is horizontal align center then
                set horizontal alignment to horizontal align general
            else
                set horizontal alignment to horizontal align center
            end if
        end if
    end tell
end tell
