option explicit

function addTextBox(rng as range, name as string) as msForms.textBox ' {

    set addTextBox = addOleObject("Forms.TextBox.1", rng).object

    addTextBox.name = name

end function ' }

function addButton(rng as range, name as string, caption as string) as msForms.commandButton ' {

     set addButton = addOleObject("Forms.CommandBUtton.1", rng).object
     addButton.name    = name
     addButton.caption = caption

end function ' }

function addOleObject(classType as string, rng as range) as oleObject ' {

     set addOleObject = rng.parent.OLEObjects.add( _
        classType       :=  classType  , _
        link            :=  false      , _
        displayAsIcon   :=  false      , _
        left            :=  rng.left   , _
        top             :=  rng.top    , _
        width           :=  rng.width  , _
        height          :=  rng.height   _
     )

end function ' }
