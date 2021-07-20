option explicit

private sub sqlText_keyDown(byVal keyCode as msForms.returnInteger, byVal state as integer) ' {

'   dim shift, alt, ctrl as string

'   if state and 1 then shift = "shf" 
'   if state and 2 then ctrl  = "ctl"
'   if state and 4 then alt   = "alt"

    if (state and 2) and (keyCode = 13) then ' {
        processQuery
    end if ' }

end sub ' }

private sub btn_click() ' {

    processQuery

end sub ' }
