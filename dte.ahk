#Requires AutoHotkey v2.0
#SingleInstance force

GlobalConfigFile := A_ScriptDir "\config.ini" ; global configuration such as username and current term
global FONT_SIZE := IniRead(GlobalConfigFile,"Appearance","FontSize")
global FONT_NAME := IniRead(GlobalConfigFile,"Appearance","Font")
global THEME := IniRead(GlobalConfigFile,"Appearance","Theme")


; **********************************************************************************************************************
; ************************************************** CLASSES SECTION ***************************************************
; **********************************************************************************************************************

class GUIBuilder {
    ; Constructor
    __New(title := "Dynamic GUI", options := "Resize"){
        this.gui := Gui(options, title)
        this.controls := Map()
        this.groupBoxes := Map()
        this.currentY := 10
        this.margin := 10
        this.controlHeight := 23
        this.controlSpacing := 5
    }

    ; Create a dynamic GroupBox with children controls
    CreateGroupBox(name, title, x := 10, y := "", width := 300, height := 200){
        if(y == ""){
            y := this.currentY
        }

        ; Create the GroupBox
        gb := this.gui.Add("GroupBox", Format("x{1} y{2} w{3} h{4}", x, y, width, height), title)

        ; Create a container to track this GroupBox's controls and positioning
        groupData := {
            groupBox: gb,
            controls: Map(),
            currentY: y + 25,
            x: x + this.margin,
            width: width - (this.margin * 2),
            rightEdge: x + width
        }

        this.groupBoxes[name] := groupData
        this.controls[name] := gb

        ; Update global Y position
        this.currentY := y + height + this.margin
    
        return groupData
    }

    ; Add Text-Edit pair to a GroupBox
    AddTextEditPair(groupName, controlName, labelText, editWidth := 150, defaultValue := ""){
        if(!this.groupBoxes.Has(groupName)){
            throw Error("GroupBox '" . groupName . "' not found.")
        }

        group := this.groupBoxes[groupName]

        ; Add a label
        labelCtrl := this.gui.Add("Text", Format("x{1} y{2} w80", group.x, group.currentY), labelText)

        ; Add edit control next to label
        editCtrl := this.gui.Add("Edit", Format("x{1} y{2} w{3}", group.x + 85, group.currentY, editWidth), defaultValue)

        ; Store controls
        group.controls[controlName] := {label: labelCtrl, edit: editCtrl}
        this.controls[controlName . "_label"] := labelCtrl
        this.controls[controlName . "_edit"] := editCtrl

        ; Update Y position for next control
        group.currentY += this.controlHeight + this.controlSpacing

        return {label: labelCtrl, edit: editCtrl}
    }

    ; Add Text-DropDownList pair to a GroupBox
    AddTextDropDownListPair(groupName, controlName, labelText, items := [], dropDownWidth := 150, defaultSelection := 1){
        if(!this.groupBoxes.Has(groupName)){
            throw Error("GroupBox '" . groupName . "' not found.")
        }

        group := this.groupBoxes[groupName]

        ; Add a label
        labelCtrl := this.gui.Add("Text", Format("x{1} y{2} w80", group.x, group.currentY), labelText)

        ; Add DropDownList control next to label
        dropDownCtrl := this.gui.Add("DropDownList", Format("x{1} y{2} w{3} Choose{4}", group.x + 85, group.currentY, dropDownWidth, defaultSelection), items)

        ; Store controls
        group.controls[controlName] := {label: labelCtrl, edit: dropDownCtrl}
        this.controls[controlName . "_label"] := labelCtrl
        this.controls[controlName . "_dropdown"] := dropDownCtrl

        ; Update Y position for next control
        group.currentY += this.controlHeight + this.controlSpacing

        return {label: labelCtrl, dropDown: dropDownCtrl}
    }

    ; Add Text-CheckBox pair to a GroupBox
    AddTextCheckBoxPair(groupName, controlName, labelText, checkBoxText := "", isChecked := false, checkBoxWidth := 150){
        if(!this.groupBoxes.Has(groupName)){
            throw Error("GroupBox '" . groupName . "' not found.")
        }

        group := this.groupBoxes[groupName]

        ; Add a label
        labelCtrl := this.gui.Add("Text", Format("x{1} y{2} w80", group.x, group.currentY), labelText)

        ; Add checkbox control next to label
        checkOptions := Format("x{1} y{2} w{3}", group.x + 85, group.currentY, checkBoxWidth)
        if (isChecked){
            checkOptions .= " Checked"
        }
        checkBoxCtrl := this.gui.Add("CheckBox", checkOptions, checkBoxText)

        ; Store controls
        group.controls[controlName] := {label: labelCtrl, edit: checkBoxCtrl}
        this.controls[controlName . "_label"] := labelCtrl
        this.controls[controlName . "_checkbox"] := checkBoxCtrl

        ; Update Y position for next control
        group.currentY += this.controlHeight + this.controlSpacing

        return {label: labelCtrl, dropDown: checkBoxCtrl}
    }

    AddUpdateButton(groupName, controlName, buttonText, buttonXpos, buttonYpos, buttonWidth := 80){
        if(!this.groupBoxes.Has(groupName)){
            throw Error("GroupBox '" . groupName . "' not found.")
        }

        group := this.groupBoxes[groupName]

        ; Add the button ; FIX: x and y positions as parameters
        buttonCtrl := this.gui.Add("Button",Format("x{1} y{2} w{3}", group.x + buttonXpos, buttonYpos, buttonWidth), buttonText)

        ; Store controls
        group.controls[controlName] := {button: buttonCtrl}
        this.controls[controlName . "_button"] := buttonCtrl

        ; Update Y position for next control
        group.currentY += this.controlHeight + this.controlSpacing

        return {button: buttonCtrl}
    }

    ; Get control value by name (works with different control types)
    GetValue(controlName){
        if(this.controls.Has(controlName . "_edit")){
            return this.controls[controlName . "_edit"].Text

        } else if(this.controls.Has(controlName . "_dropdown")){
            return this.controls[controlName . "_dropdown"].Text

        } else if(this.controls.Has(controlName . "_checkbox")){
            return this.controls[controlName . "_checkbox"].Value

        } else if(this.controls.Has(controlName)){
            return this.controls[controlName].Text

        } else {
            throw Error("Control '" . controlName . "' not found.")
        }
    }

    ; Set control value by name
    SetValue(controlName, value){
        if(this.controls.Has(controlName . "_edit")){
            this.controls[controlName . "_edit"].Text := value

        } else if(this.controls.Has(controlName . "_dropdown")){
            this.controls[controlName . "_dropdown"].Choose(value)

        } else if(this.controls.Has(controlName . "_checkbox")){
            this.controls[controlName . "_checkbox"].Value := value

        } else if(this.controls.Has(controlName)){
            this.controls[controlName].Text := value

        } else {
            throw Error("Control '" . controlName . "' not found.")
        }
    }

    ; Show the GUI
    Show(options := "w400 h300") {
        this.gui.Show(options)
    }

    ; Add standard buttons (OK, Cancel, Apply, etc.)
    AddStandardButtons(buttons := ["OK", "Cancel"]){
        buttonWidth := 75
        buttonSpacing := 10
        totalWidth := (buttonWidth * buttons.Length) + (buttonSpacing * (buttons.Length - 1))

        ; Calculate starting X position to center buttons
        startX := (400 - totalWidth) / 2

        currentX := startX
        buttonControls := Map()

        for index, buttonText in buttons {
            btn := this.gui.Add("Button", Format("x{1} y{2} w{3} h25", currentX, this.currentY, buttonWidth), buttonText)
            buttonControls[buttonText] := btn
            this.controls[buttonText . "_button"] := btn
            currentX += buttonWidth + buttonSpacing
        }

        this.currentY += 35
        return buttonControls
    }
}


; **********************************************************************************************************************
; ************************************************** FUNCTION SECTION **************************************************
; **********************************************************************************************************************

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& UTILITIES &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

; A utility function for loading the magic strings when main functions are called.
; Important in case the user doesn't have excel or chrome open when they initially
; start up the script
UpdateGlobalStrings(){
    try {
        ; Magic Strings that are re-used throughout
        Global BANNER_APP := "ahk_exe chrome.exe"
        Global EXCEL_APP := ComObjActive("Excel.Application")
    } catch Error {
        MsgBox("At least one required application is not open.","DTE Helper Tool","T2")
    }
}

UpdateLayoutSpacing(){
    global Layout := Map()

    Layout["fontSize"] := FONT_SIZE
    Layout["marginX"] := Ceil(FONT_SIZE * 1.25)
    Layout["marginY"] := Ceil(FONT_SIZE * 0.75)

    Layout["smallGap"] := Ceil(FONT_SIZE * 0.6)
    Layout["mediumGap"] := Ceil(FONT_SIZE * 1.2)
    Layout["largeGap"] := Ceil(FONT_SIZE * 1.8)
}
UpdateLayoutSpacing()

; A utility function to display the current global configurations
ViewConfig(){
    SpeedContext := ""
    If(Integer(IniRead(GlobalConfigFile,"SessionContext","SystemSpeed")) = 1){
        SpeedContext := "Normal"
    } Else If(Integer(IniRead(GlobalConfigFile,"SessionContext","SystemSpeed")) > 5){
        SpeedContext := "Very Slow"
    } Else {
        SpeedContext := "Slow"
    }

    MsgBox(
        "Username: " . IniRead(GlobalConfigFile,"Local","User") . "`n" .
        "Current Term: " . IniRead(GlobalConfigFile,"TermContext","CurrentTerm") . "`n" .
        "System Speed: " . IniRead(GlobalConfigFile,"SessionContext","SystemSpeed") . " (" . SpeedContext . ")"  . "`n" .
        "Theme: " . IniRead(GlobalConfigFile,"Appearance","Theme") . "`n" .
        "Font: " . FONT_NAME . "`n" .
        "Font Size: " . IniRead(GlobalConfigFile,"Appearance","FontSize"),
        "DTE  |  Global Configurations",
        4096
    )
}

; A utility function for updating global configurations, such as the username during setup, or the current term
UpdateConfig(GuiControlObj,Info,Section,Key,editControl){
    if(Section != "TermContext" && Section != "Appearance"){
        local newValue := editControl.Value
        IniWrite(newValue,GlobalConfigFile,Section,Key)
        MsgBox("Global setting: '" . Key . "' updated to: " . newValue, "Success", 4096)
        Reload

    ; Updating the current term auto-updates the previous and next term codes in config settings
    ; the script will use these config values, together with the TermSplitter() function, the insert
    ; term messages and codes as needed during DTE tasks
    } else if(Section = "TermContext" && Section != "Appearance") {
        ; this is the new term code
        local newValue := editControl.Value
        ; for fall and winter term codes, generate next and previous by adding/subtracting 10
        if(SubStr(String(newValue),5,1) = "2" || SubStr(String(newValue),5,1) = "3"){
            local pTerm := newValue - 10
            local nTerm := newValue + 10
            IniWrite(newValue,GlobalConfigFile,Section,Key)
            IniWrite(pTerm,GlobalConfigFile,"TermContext","PreviousTerm")
            IniWrite(nTerm,GlobalConfigFile,"TermContext","NextTerm")
        ; for summer term codes, generate next by adding 10, generate previous by subtracting 70
        } else if(SubStr(String(newValue),5,1) = "1") {
            local pTerm := newValue - 70
            local nTerm := newValue + 10
            IniWrite(newValue,GlobalConfigFile,Section,Key)
            IniWrite(pTerm,GlobalConfigFile,"TermContext","PreviousTerm")
            IniWrite(nTerm,GlobalConfigFile,"TermContext","NextTerm")
        ; for spring term codes, generate next by adding 70, generate previous by subtracting 10
        } else if(SubStr(String(newValue),5,1) = "4") {
            local pTerm := newValue - 10
            local nTerm := newValue + 70
            IniWrite(newValue,GlobalConfigFile,Section,Key)
            IniWrite(pTerm,GlobalConfigFile,"TermContext","PreviousTerm")
            IniWrite(nTerm,GlobalConfigFile,"TermContext","NextTerm")
        } else {
            MsgBox("Error when attempting to update Term Context configs.`n See Software Admin","Error")
        }

        MsgBox("Global setting: '" . Key . "' updated to: " . newValue, "Success", 4096)
        Reload
    } else if(Section = "Appearance" && Section != "TermContext") {
        local newValue := editControl.Text
        IniWrite(newValue,GlobalConfigFile,Section,Key)
        MsgBox("Global setting: '" . Key . "' updated to: " . newValue, "Success", 4096)
        Reload
    } else {
        MsgBox("Error when attempting to update global configs.`n See Software Admin","Error")
    }
}

; A utility function that 'wraps' keystrokes in pauses/sleeps
Press(btn,repeats,preSleep,postSleep){
    global GlobalConfigFile
    if(repeats != 1){
        Loop repeats {
            Sleep preSleep * Integer(IniRead(GlobalConfigFile,"SessionContext","SystemSpeed",))
            Send btn
            Sleep postSleep * Integer(IniRead(GlobalConfigFile,"SessionContext","SystemSpeed",))
        }
    } else {
        Sleep preSleep * Integer(IniRead(GlobalConfigFile,"SessionContext","SystemSpeed",))
        Send btn
        Sleep postSleep * Integer(IniRead(GlobalConfigFile,"SessionContext","SystemSpeed",))
    }
}

; A utility function for transforming LCC's term codes into longform term descriptions
TermSplitter(Term){
    IsoYear := Number(SubStr(Term, 1, 4))
    IsoTerm := Number(SubStr(Term, -2))

    If(IsoTerm = 10){
        TermTitle := "Summer"
        YearTitle := IsoYear - 1
    } Else If(IsoTerm = 20){
        TermTitle := "Fall"
        YearTitle := IsoYear - 1
    } Else If(IsoTerm = 30){
        TermTitle := "Winter"
        YearTitle := IsoYear
    } Else If(IsoTerm = 40){
        TermTitle := "Spring"
        YearTitle := IsoYear
    }
    LongTerm := TermTitle . " " . YearTitle
    return LongTerm
}

; A utility function for handling the various snippets that get used when processing program changes
SendTermSnippet(snippet){
    ; Replace placeholders like {NextTerm} with actual term
    nextTerm := TermSplitter(IniRead(GlobalConfigFile, "TermContext", "NextTerm"))
    currentTerm := TermSplitter(IniRead(GlobalConfigFile, "TermContext", "CurrentTerm"))

    snippet := StrReplace(snippet, "{NextTerm}", nextTerm)
    snippet := StrReplace(snippet, "{CurrentTerm}", currentTerm)

    Send snippet
}

; A utility function that handles toggling checkboxes for attributes in Articulator
ChangeCheck(GuiControlObj, Info,ArrayOfAttr,Attr){
    If(GuiControlObj.Value){
        ArrayOfAttr.Push(Attr)
        MsgBox("Attribute added. Current array: " . ArrayOfAttr.Join(" ,"))
    } else {

    }
}

; A utility function for marking areas as under development
DevMsg(){
    MsgBox("This feature is under development","DTE Helper Tool","T2.5")
}

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& General Use Callbacks &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; A callback function for auto-tabbing in edit fields that have a max number of characters
AutoTabGUI(Ctrl, Info, Chars){
    EditContent := Ctrl.Value
    If(StrLen(EditContent) >= Chars){
        Send "{Tab}"
    }
}

; A callback function for closing GUIs on cancel click by reloading the script
hardCloseGUI(GuiControlObj, Info){
    ; MsgBox("You clicked cancel `n The script will restart.","DTE Helper Tool","T1.5")
    Reload
}

; A callback function for clicking the OK button in a GUI, when OK doesn't do anything
OKClickBland(builderRef, *){
    builderRef.gui.Hide()
}

; A callback function for clicking the cancel button in a GUI
CancelClick(builderRef, *){
    builderRef.gui.Hide()
}

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& TERM BUILDER Callbacks &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; A callback function for moving from the main GUI to the TermBuilder GUI
tbOPENbtn(GuiControlObj, Info){
    TermBuilder_GUI.Show("AutoSize Center")
    dteGui.Hide()
}

; A callback function for closing the Term Builder GUI and returning to the main GUI
tbToMain(GuiControlObj, Info){
    TermBuilder_GUI.Hide()
    dteGui.Restore()
}

; A callback function to initiate term creation from the TermBuilder GUI
tbGObtn(GuiControlObj, Info, InstNum, FICE, startRow, endRow){
    TermBuilder_GUI.Hide()
    TermBuilder(instNum_tbGUI,FICE_tbgui,start_tbgui,end_tbgui)
}

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& TRANSCRIPT REVIEW Callbacks &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; A callback function for moving from the main GUI to the TranscriptReview GUI
trOPENbtn(GuiControlObj, Info){
    TranscriptReview_GUI.Show("AutoSize Center")
    dteGui.Hide()
}

; A callback function& for closing the Transcript Review GUI and returning to the main GUI
trToMain(GuiControlObj, Info){
    TranscriptReview_GUI.Hide()
    dteGui.Restore()
}

; A callback function to initiate transcript review from the TranscriptReview GUI
trGObtn(GuiControlObj, Info, startRow, endRow){
    TranscriptReview_GUI.Hide()
    TranscriptReview(startRow,endRow)
}

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& PROGRAM CHANGES Callbacks &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; A callback function for moving from the main GUI to the ProgramChange GUI
posOPENbtn(GuiControlObj, Info){
    ProgramChange_GUI.Show("AutoSize Center")
    dteGui.Hide()
}

posToMain(GuiControlObj, Info){
    ProgramChange_GUI.Hide()
    dteGui.Restore()
}

posGObtn(GuiControlObj, Info){
    ProgramChange_GUI.Hide()
    DevMsg
    dteGui.Show()
}

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& DIPLOMA DATES Callbacks &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; A callback function for moving from the main GUI to the DiplomaDates GUI
diplOPENbtn(GuiControlObj, Info){
    DiplDates_GUI.Show("AutoSize Center")
    dteGui.Hide()
}

diplToMain(GuiControlObj, Info){
    DiplDates_GUI.Hide()
    dteGui.Restore()
}

; A callback function to initiate the mass batch diploma date process
diplGObtn(GuiControlObj, Info, startRow, endRow, orderedOn, mailedOn){
    DiplomaDates(startRow,endRow,orderedOn,mailedOn)
}

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& ARTICULATOR Callbacks &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; A callback function for moving from the main GUI to the Articulator GUI
articOPENbtn(GuiControlObj, Info){
    Articulator_GUI.Show("AutoSize Center")
    dteGui.Hide()
}

; A callback function for closing the Articulator GUI and returning to the main GUI
articToMain(GuiControlObj, Info){
    Articulator_GUI.Hide()
    dteGui.Restore()
}

; A callback function to initiate the articulator function
articGObtn(GuiControlObj, Info, startRow, endRow){
    Articulator_GUI.Hide()
    Articulator(startRow, endRow)
}

; A callback for turning the map/menu of attributes into an array of the actual attributes we're sending
articPusher(GuiControlObj, Info, MapOfAttr, Attribute){
    MapOfAttr[Attribute] := MapOfAttr[Attribute] * -1
    
}

; A callback function for adding attributes to equivalent LCC courses
articAddAttr(GuiControlObj, Info, Attributes){
    If (WinExist(BANNER_APP)){
        WinActivate
    }

    sendingAttr := []
    For(Key, Value in Attributes){
        If(Integer(Attributes[Key]) = 1){
            sendingAttr.Push(Key)
        }
    }
    
    Attributes_GUI.Hide()

    counter := 1
    If(sendingAttr.Length >= 1){
        While(counter < sendingAttr.Length){
            Press(sendingAttr[counter],1,200,200)
            Press("{down}",1,200,200)
            counter++
        }
        Press(sendingAttr[sendingAttr.Length],1,200,200)
        Press("{F10}",1,200,200)
    } else {
        Press("{F10}",1,200,200)
    }
        Pause(-1)
}

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& DTE PROCESSES &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

; A function for processing mass batches of order & mail dates for outgoing diplomas
; SHADIPL
DiplomaDates(startRow, endRow, orderDate, mailDate){
    UpdateGlobalStrings()
    DiplomaDates_excel := EXCEL_APP
    if (!DiplomaDates_excel) {
        MsgBox("Excel is not open.",,"T2")
    } else {
        ; grab the Banner window
        If (WinExist(BANNER_APP)){
                WinActivate
            }
        
        ;define the active row first by the start row then loop through as many times as needed to get to the end row
        activeRow := startRow.Value

        ;define the columns for our data/variables
        IDCol := "A"
        DegSeqCol := "B"

        While (activeRow <= endRow.Value){
            ; define the cells for the active row
            IDCell := IDCol . activeRow
            DegSeqCell := DegSeqCol . activeRow

            ; grab values for a given row from the excel sheet
            ID := DiplomaDates_excel.Range(IDCell).Text
            DegSeq := DiplomaDates_excel.Range(DegSeqCell).Text

            ; grab the SHATAEQ window
            WinActivate BANNER_APP
            
            Press(ID,1,200,200)

            Press(DegSeq,1,200,200)

            Press("!{PgDn}",1,100,200)
            Press("{Esc}",1,0,200)

            Press("{Tab}",5,200,200)

            ; ORDER DATE
            Press(orderDate,1,200,200)

            Press("{Tab}",1,200,200)

            ; MAILED DATE
            Press(mailDate,1,200,200)

            Press("{F10}",1,200,200)

            Press("{F5}",1,200,200)

            WinActivate "Excel ahk_class XLMAIN"

            Press("^{b}",1,200,200)

            Press("{Down}",1,200,200)

            ;prep next iteration of the loop
            activeRow++
            }
            return
    }
}

; A set of functions for creating transfer terms as part of the degree evaluation process
; SHATRNS
TermBuilder(InstNum, FICE, startRow, endRow){
    UpdateGlobalStrings()
    ;Data Entry Function
    TermBuilder_DataEntry(InstNum, FICE, prevTermSeqNo, TermSeqNo, Term, LongTerm){
        If(TermSeqNo != prevTermSeqNo){
            ; handle non-first terms, i.e. skip making the transfer institution
            If(TermSeqNo != 1){

                Press("{Tab}",2,200,200)

                Press(TermSeqNo,1,200,200)

                Loop 2 {
                    Press("!{PgDn}",1,200,200)
                    Press("{Esc}",1,200,200)
                }

                Press(LongTerm,1,200,200)

                Press("{Tab}",2,200,200)
                
                Press(Term,1,200,200)

                Press("{Tab}",2,200,200)

                Press("CR",1,200,200)

                Press("{F10}",2,200,200)   

                Press("{F5}",2,200,200)
                
            } Else {
                ; make the transfer institution for the first term sequence
                Press("{Tab}",1,200,200)

                Press(InstNum,1,200,200)

                Press("{Tab}",1,200,200)

                Press(TermSeqNo,1,200,200)

                Press("!{PgDn}",1,200,200)

                Press("{Esc}",1,200,200)

                Press(FICE,1,200,200)

                Press("{Tab}",2,200,200)

                Press("{Space}",1,200,200)

                Press("{F10}",1,200,200)

                Press("!{PgDn}",1,200,200)

                Press("{Esc}",1,200,200)

                Press(LongTerm,1,200,200)

                Press("{Tab}",2,200,200)

                Press(Term,1,200,200)

                Press("{Tab}",2,200,200)

                Press("CR",1,200,200)

                Press("{F10}",1,200,200)

                Press("{F5}",1,200,200)
            }
        } Else {
            return
        }
        return
    }

    ; Loop to grab data that is then input using the data entry function
    ; initiate excel connection, let the user know if excel is not open
    TermBuilder_excel := EXCEL_APP
    if (!TermBuilder_excel) {
        MsgBox("Excel is not open.",,"T2")
    } else {
        ; proceed normally if excel is open    
        ; grab the Banner window
        If (WinExist(BANNER_APP)){
            WinActivate
        } else {
            MsgBox("Chrome is not open.",,"T2")
        }
    
        ;define the active row first by the start row then loop through as many times as needed to get to the end row
        activeRow := startRow.Value
        prevTermSeqNo := ""

        ;define the columns for our data/variables
        TermSeqCol := "A"
        TermCol := "B"

        While (activeRow <= endRow.Value) {
            ; define the cells for the active row
            TermSeqCell := TermSeqCol . activeRow
            TermCell := TermCol . activeRow

            ;grab values for a given row from the excel sheet
            TermSeqNo := Floor(TermBuilder_excel.Range(TermSeqCell).Value)
            Term := TermBuilder_excel.Range(TermCell).Text

            ; transform term code into term description (LongTerm, e.g., Fall 2024)
            LongTerm := TermSplitter(Term)

            ; grab the Banner window
            If (WinExist(BANNER_APP)){
                WinActivate
            }

            ; call the function
            TermBuilder_DataEntry(InstNum.Value,FICE.Value,prevTermSeqNo,TermSeqNo,Term,LongTerm)

            ; prep next iteration of the loop
            activeRow++
            prevTermSeqNo := TermSeqNo
        }
        return
    }
}

; A set of functions for entering transfer course data into SHATAEQ as part of the degree evaluation process
; SHATAEQ
TranscriptReview(startRow, endRow){
    UpdateGlobalStrings()
    ; Data entry function
    TranscriptReview_DataEntry(TransferCourseData,prevTermSeqNo){
        ; Handle term seq numbers with single digits, and double digits
        If (TransferCourseData["TermSeqNo"] < 10) {
            Press(TransferCourseData["TermSeqNo"],1,200,200)
            Press("{Tab}",1,200,200)
        } Else {
            Press(TransferCourseData["TermSeqNo"],1,200,200)
        }
        ; handle same term seq number as previous
        If (TransferCourseData["TermSeqNo"] = prevTermSeqNo) {
            Press("{Tab}",3,200,200)

            Press(TransferCourseData["Subj"],1,200,200)

            Press("{Tab}",1,200,200)

            Press(TransferCourseData["Course"],1,200,200)

            Press("{Tab}",1,200,200)

            Press(TransferCourseData["Credits"],1,200,200)

            Press("{Tab}",1,200,200)

            Press("{Raw}" . TransferCourseData["Grade"],1,200,200)

            Press("{Tab}",1,200,200)

            Press(TransferCourseData["Duplicate"],1,200,200)
        } Else {
            ; normal case for first entry in a given term seq number
            Press(TransferCourseData["Term"],1,200,200)
            
            Press("CR",1,200,200)

            Press("{Tab}",1,200,200)
            
            Press(TransferCourseData["Subj"],1,200,200)

            Press("{Tab}",1,200,200)

            Press(TransferCourseData["Course"],1,200,200)

            Press("{Tab}",1,200,200)

            Press(TransferCourseData["Credits"],1,200,200)

            Press("{Tab}",1,200,200)

            Press("{Raw}" . TransferCourseData["Grade"],1,200,200)

            Press("{Tab}",1,200,200)

            Press(TransferCourseData["Duplicate"],1,200,200)
        }

        ; handle override edit titles
        If (TransferCourseData["CustomTitle"] != ""){
            Press("{Tab}",1,200,200)
            
            Press(TransferCourseData["CustomTitle"],1,200,200)

            Press("{F10}",1,200,200)

            Press("{F3}",1,200,200)

            Press("{F10}",1,200,200)

            Press("{Tab}",10,200,200)

            Press("Over",1,200,200)

            Press("{Tab}",9,200,200)

            Press(TransferCourseData["CustomTitle"],1,200,200)

            Press("{F10}",1,200,200)

            Press("{Down}",1,200,200)

        } Else {
            Press("{Down}",1,200,200)
        }
    }

    ; initiate excel connection
    TranscriptReview_excel := EXCEL_APP
    if (!TranscriptReview_excel) {
            MsgBox("Excel is not open.",,"T2")
        } else {
            ; grab the Banner window
            If (WinExist(BANNER_APP)){
                    WinActivate
                }
            
            ;define the active row first by the start row then loop through as many times as needed to get to the end row
            activeRow := startRow.Value
            prevTermSeqNo := ""

            ;define the columns for our data/variables
            TermSeqCol := "A"
            TermCol := "B"
            SubjCol := "C"
            CourseCol := "D"
            CreditsCol := "F"
            GradeCol := "G"
            DupCol := "H"
            CustomTitleCol := "I"

            ; run through the loop, from the user start
            ; to the user end rows
            While (activeRow <= endRow.Value){
                ; define the cells for the active row
                TermSeqCell := TermSeqCol . activeRow
                TermCell := TermCol . activeRow
                SubjCell := SubjCol . activeRow
                CourseCell := CourseCol . activeRow
                CreditsCell := CreditsCol . activeRow
                GradeCell := GradeCol . activeRow
                DupCell := DupCol . activeRow
                CustomTitleCell := CustomTitleCol . activeRow

                TransferCourseData := Map()

                ; grab values for a given row from the excel sheet
                TransferCourseData["TermSeqNo"] := Floor(TranscriptReview_excel.Range(TermSeqCell).Value)
                TransferCourseData["Term"] := TranscriptReview_excel.Range(TermCell).Text
                TransferCourseData["Subj"] := TranscriptReview_excel.Range(SubjCell).Text
                TransferCourseData["Course"] := TranscriptReview_excel.Range(CourseCell).Text
                TransferCourseData["Credits"] := TranscriptReview_excel.Range(CreditsCell).Text
                TransferCourseData["Grade"] := TranscriptReview_excel.Range(GradeCell).Value 
                TransferCourseData["Duplicate"] := TranscriptReview_excel.Range(DupCell).Text
                TransferCourseData["CustomTitle"] := TranscriptReview_excel.Range(CustomTitleCell).Text

                ; grab the SHATAEQ window
                WinActivate BANNER_APP
                
                ; start plugging and chugging
                TranscriptReview_DataEntry(TransferCourseData, prevTermSeqNo)

                ;prep next iteration of the loop
                activeRow++
                prevTermSeqNo .= TransferCourseData["TermSeqNo"]
            }
            return
        }
}

; A function for adding new course articulations, or updating existing ones, to the transfer table
; SHATATR
Articulator(startRow,endRow){
    UpdateGlobalStrings()
    ; use global config to generate the catalog year and first bit of the save message
    artCatYear := SubStr(IniRead(GlobalConfigFile,"TermContext","CurrentTerm"),1,4)
    artMsgStart := IniRead(GlobalConfigFile,"TermContext","CurrentTerm") . " " . StrUpper(SubStr(IniRead(GlobalConfigFile,"Local","User"),-1,1)) . StrUpper(SubStr(IniRead(GlobalConfigFile,"Local","User"),1,1))
    
    ; Date Entry Function
    Articulator_DataEntry(transferCourse,equivalentCourse,action,message,catalogYear){
        If (WinExist(BANNER_APP)){
            WinActivate
        }

        ; proceed based on case (No Change, Add, Update)
        if(action = "No Change") {
            return
        } else if(action = "Add") {
            ; If ADD action, then create a new line and proceed as the default case
            ; Add a new line in SHATATR
            Press("{F6}",1,200,200)

            ; Tab to subject field
            Press("{Tab}",2,200,200)

            ; Send transfer subject code
            Press(transferCourse["Subject"],1,200,200)

            ; Tab to number field
            Press("{Tab}",1,200,200)

            ; Send transfer number
            Press(transferCourse["Course"],1,200,200)

            ; Tab to title field
            Press("{Tab}",1,200,200)

            ; Send the transfer title, and either tab to next field, or let banner auto-tab
            ; if the title is 30 characters (max for this field)
            If (StrLen(transferCourse["Title"]) < 30) {
                Press(transferCourse["Title"],1,200,200)
                Press("{Tab}",1,200,200)
            } Else If (StrLen(transferCourse["Title"]) = 30) {
                Press(transferCourse["Title"],1,200,200)
            }

            ; Send effective start term
            Press(transferCourse["Term"],1,200,200)

            ; Tab to equivalent exists field
            Press("{Tab}",1,200,200)

            ; Send "yes"
            Press("Yes",1,200,200)

            ; Tab to catalog year field
            Press("{Tab}",1,200,200)

            ; Send catalog year
            Press(catalogYear,1,200,200)

            ; Send CR level
            Press("CR",1,200,200)

            ; Tab to status field
            Press("{Tab}",1,200,200)

            ; Send AC (active)
            Press("AC",1,200,200)

            ; Tab to credits low field
            Press("{Tab}",2,200,200)

            ; Send credits low
            Press(transferCourse["Credits"],1,200,200)

            ; SAVE
            Press("{F10}",1,200,200)

            ; Page down twice to the equivalent course table
            Loop 2 {
                Press("!{PgDn}",1,300,500)
                Press("{Esc}",1,400,400)
            }

            ; Send equivalent subject
            Press(equivalentCourse["Subject"],1,200,200)
            Press("{Tab}",1,200,200)

            ; Send equivalent course
            Press(equivalentCourse["Course"],1,200,200)
            Press("{Tab}",1,200,200)

            ; Send equivalent title
            Press(equivalentCourse["Title"],1,200,200)

            ; SAVE
            Press("{F10}",1,200,200)

            ; Page down to the equivalent course attributes table
            Press("!{PgDn}",1,400,400)
            Press("{Esc}",1,400,400)

            ; Delete existing attributes
            Loop 5 {
                Press("+{F6}",1,200,200)
            }
            ; SAVE
            Press("{F10}",1,200,200)

            ; SEND ANY ATTRIBUTES
            Attributes_GUI.Show("AutoSize Center")
            Pause

            ; Grab the banner window
            If (WinExist(BANNER_APP)){
                WinActivate
            }

            ; Page down to the equivalent course comments table
            Press("!{PgDn}",1,400,400)
            Press("{Esc}",1,400,400)

            ; Send save message
            Press(message,1,200,200)

            ; SAVE and return to start
            Press("{F10}",1,200,200)
            Press("!{PgDn}",1,400,400)
            Press("{Esc}",1,400,400)

        } else if(action = "Update") {
            ; if UPDATE action then proceed with filtering for the existing course, then update
            ; as needed
            ; filter to the course being updated
            Press("{F7}",1,200,800)
            Press("{Tab}",8,200,200)
            Press(transferCourse["Subject"],1,200,200)
            Press("{Tab}",2,300,300)
            Press(transferCourse["Course"],1,200,200)
            Press("{Tab}",4,200,200)
            Press("{Down}",1,200,200)
            Loop 2 {
                Press("{Tab}",2,200,200)
                Press("{Down}",1,200,200)
            }
            Press(transferCourse["Term"],1,200,200)
            Press("{Enter}",1,200,500)

            ; update catalog year
            Press("{Tab}",7,200,200)
            Press(catalogYear,1,200,200)

            ; page down to equivalent course that is being updated
            Loop 2 {
                Press("!{PgDn}",1,400,400)
                Press("{Esc}",1,400,400)
            }
            ; delete existing course
            Press("+{F6}",1,500,200)

            ; Send equivalent subject
            Press(equivalentCourse["Subject"],1,200,200)
            Press("{Tab}",1,200,200)

            ; Send equivalent course
            Press(equivalentCourse["Course"],1,200,200)
            Press("{Tab}",1,200,200)

            ; Send equivalent title
            Press(equivalentCourse["Title"],1,200,200)
            Press("{Tab}",1,200,200)

            ; SAVE
            Press("{F10}",1,200,200)

            ; Page down to the equivalent course attributes table
            Press("!{PgDn}",1,400,400)
            Press("{Esc}",1,400,400)

            ; Delete existing attributes
            Loop 5 {
                Press("+{F6}",1,200,200)
            }

            ; SEND ANY ATTRIBUTES
            Attributes_GUI.Show("AutoSize Center")
            Pause

            ; Grab the Banner window
            If (WinExist(BANNER_APP)){
                WinActivate
            }

            ; Page down to the equivalent course comments table
            Press("!{PgDn}",1,400,400)
            Press("{Esc}",1,400,400)

            ; Send save message
            Press(message,1,200,200)

            ; SAVE and return to start, clearing the filter
            Press("{F10}",1,200,200)
            Press("!{PgDn}",1,400,400)
            Press("{Esc}",1,400,400)
            Press("{F7}",1,300,200)
            Press("{Esc}",1,400,2000)
        } else {
            return
        }
    } ; END OF DATA ENTRY FUNCTION

    ; initiate excel connection
    Articulator_excel := EXCEL_APP
    if (!Articulator_excel) {
            MsgBox("Excel is not open.",,"T2")
        } else {
            ; grab the Banner window
            If (WinExist(BANNER_APP)){
                    WinActivate
                }
            
            ; define the active row first by the start row then loop through as many times as needed to get to the end row
            activeRow := startRow.Value ; needs .Value when called from GUI
            
            ; define the columns for our data/variables
            ; Transfer data
            TRSubjCol := "C"
            TRCourseCol := "D"
            TRTitleCol := "E"
            TRTermCol := "M"
            TRCreditsCol := "F"
            ; equiv LCC data            
            LCCSubjCol := "J"
            LCCCourseCol := "K"
            LCCTitleCol := "L"
            LCCCreditsCol := "N"

            ActionCol := "O"
            ComboCol := "P"

            ; run through the loop, from the user start
            ; to the user end rows
            While (activeRow <= endRow.Value){
                ; define the transfer course cells for the active row
                TRSubjCell := TRSubjCol . activeRow
                TRCourseCell := TRCourseCol . activeRow
                TRTitleCell := TRTitleCol . activeRow
                TRTermCell := TRTermCol . activeRow
                TRCreditsCell := TRCreditsCol . activeRow

                LCCSubjCell := LCCSubjCol . activeRow
                LCCCourseCell := LCCCourseCol . activeRow
                LCCTitleCell := LCCTitleCol . activeRow
                LCCCreditsCell := LCCCreditsCol . activeRow

                ActionCell := ActionCol . activeRow
                ComboCell := ComboCol . activeRow
                CalendarCell := "Q2"

                ; grab values for a given row from the excel sheet
                ; transfer course data map
                transferCourse := Map()
                transferCourse["Subject"] := Articulator_excel.Range(TRSubjCell).Text
                transferCourse["Course"] := Articulator_excel.Range(TRCourseCell).Text
                transferCourse["Title"] := Articulator_excel.Range(TRTitleCell).Text
                transferCourse["Term"] := Articulator_excel.Range(TRTermCell).Text
                transferCourse["Credits"] := Articulator_excel.Range(TRCreditsCell).Text
                transferCourse["Combo"] := Articulator_excel.Range(ComboCell).Text
                transferCourse["Calendar"] := Articulator_excel.Range(CalendarCell).Text
                ; equivalent course data map
                equivalentCourse := Map()
                equivalentCourse["Subject"] := Articulator_excel.Range(LCCSubjCell).Text
                equivalentCourse["Course"] := Articulator_excel.Range(LCCCourseCell).Text
                equivalentCourse["Title"] := Articulator_excel.Range(LCCTitleCell).Text
                equivalentCourse["Credits"] := Articulator_excel.Range(LCCCreditsCell).Text

                SHATATRaction := Articulator_excel.Range(ActionCell).Text
                SHATATRmessage := artMsgStart . " " . SubStr(transferCourse["Term"],1,4) . "-" . artCatYear

                ; grab the SHATAEQ window
                WinActivate BANNER_APP
                
                ; start plugging and chugging
                Articulator_DataEntry(transferCourse,equivalentCourse,SHATATRaction,SHATATRmessage,artCatYear)

                ;prep next iteration of the loop
                activeRow++
            }
            return
    }
}


; **********************************************************************************************************************
; **************************************************** GUI SECTION *****************************************************
; **********************************************************************************************************************
; The utility GUIs and Menus

; New Look Update Global Configs GUI
UpdateConfigs(){
     ; Create a new GUI builder
    builder := GUIBuilder("DTE  |  Update Global Configurations", "Resize")
    builder.gui.SetFont("S" . FONT_SIZE, FONT_NAME)

    ; Create a GroupBox for local configuration information
    localGroup := builder.CreateGroupBox("localInfo", "Local Settings", 10, 10, 350, 69)

    ; Add controls to the GroupBox
    builder.AddTextEditPair("localInfo", "username", "Username:", 100, IniRead(GlobalConfigFile,"Local","User"))
    builder.AddUpdateButton("localInfo", "usernameUpdate", "Update", 200, "p")
    builder.controls["usernameUpdate_button"].OnEvent("Click",UpdateConfig.Bind(,,"Local","User",builder.controls["username_edit"]))

    ; Create a GroupBox for term context configuration information
    termGroup := builder.CreateGroupBox("termInfo", "Term Settings", 10, , 350, 69)

    ; Add controls to the GroupBox
    builder.AddTextEditPair("termInfo", "currentTerm", "Current Term:", 72, IniRead(GlobalConfigFile,"TermContext","CurrentTerm"))
    builder.AddUpdateButton("termInfo", "currentTermUpdate", "Update", 200, "p")
    builder.controls["currentTermUpdate_button"].OnEvent("Click",UpdateConfig.Bind(,,"TermContext","CurrentTerm",builder.controls["currentTerm_edit"]))

    ; Create a GroupBox for session configuration information
    sessionGroup := builder.CreateGroupBox("sessionInfo", "Session Settings", 10, , 350, 69)

    ; Add controls to the GroupBox
    builder.AddTextDropDownListPair("sessionInfo", "systemSpeed", "System Speed:", ["1 (Normal)","2","3","4","5 (Slow)"], 100, IniRead(GlobalConfigFile,"SessionContext","SystemSpeed"))
    builder.AddUpdateButton("sessionInfo", "systemSpeedUpdate", "Update", 200, "p")

    ; Create a GroupBox for appearance configuration information
    appearanceGroup := builder.CreateGroupBox("appearanceInfo", "Appearance Settings", 10, , 350, 121)

    ; Add controls to the GroupBox
    builder.AddTextDropDownListPair("appearanceInfo", "fontName", "Font:", ["Lucia Sans"], 200, 1)
    builder.AddTextDropDownListPair("appearanceInfo", "fontSize", "Font Size:", ["8","10","12","14","16"], 200, 3)
    builder.AddTextDropDownListPair("appearanceInfo", "theme", "Theme:", ["Light", "Dark"], 200, 1)

    ; Add standard buttons
    buttons := builder.AddStandardButtons(["OK", "Cancel"])

    ; Set up button events
    buttons["OK"].OnEvent("Click", OKClickBland.Bind(builder))
    buttons["Cancel"].OnEvent("Click", CancelClick.Bind(builder))
    
    ; Show the GUI
    builder.Show("AutoSize Center")
    
    return builder
}

; The sub-GUIs
; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& The Diploma Dates GUI &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; SHADIPL
DiplDates_GUI := Gui("AlwaysOnTop", "DTE  |  Diploma Dates")
    DiplDates_GUI.SetFont("S" . FONT_SIZE, FONT_NAME)
    DiplDates_GUI.Add("Text",,"Welcome to the diploma dates tool. `n")

    DiplDates_GUI.Add("Text","section w95   ","Order date:")
    diplOrderDate := DiplDates_GUI.Add("DateTime","w120 yp")
    orderYear := SubStr(diplOrderDate.Value,1,4)
    orderMonth := SubStr(diplOrderDate.Value,5,2)
    orderDay := SubStr(diplOrderDate.Value,7,2)
    orderedOn := orderYear "/" orderMonth "/" orderDay

    DiplDates_GUI.Add("Text","w95 xs yp+32","Mailed date:")
    diplMailDate := DiplDates_GUI.Add("DateTime","w120 yp")
    mailYear := SubStr(diplMailDate.Value,1,4)
    mailMonth := SubStr(diplMailDate.Value,5,2)
    mailDay := SubStr(diplMailDate.Value,7,2)
    mailedOn := mailYear "/" mailMonth "/" mailDay

    DiplDates_GUI.Add("Text","section ys","Start row:")
    start_dipldates := DiplDates_GUI.Add("Edit","w60 yp Number Limit 3")
    start_dipldates.OnEvent("Change",AutoTabGUI.Bind(,,3))

    DiplDates_GUI.Add("Text","xs yp+32","End row:")
    end_dipldates := DiplDates_GUI.Add("Edit","w66 yp Number Limit 3")
    end_dipldates.OnEvent("Change",AutoTabGUI.Bind(,,3))

    diplCLOSE := DiplDates_GUI.Add("Button","NoTab","Home",)
    diplCLOSE.OnEvent("Click", diplToMain)
    
    diplGO := DiplDates_GUI.Add("Button","yp","GO")
    diplGO.OnEvent("Click",diplGObtn.Bind(,,start_dipldates,end_dipldates,orderedOn,mailedOn))

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& The Term Builder GUI &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; SHATRNS
TermBuilder_GUI := Gui("AlwaysOnTop", "DTE  |  Term Builder")
    TermBuilder_GUI.SetFont("S" . FONT_SIZE, FONT_NAME)
    TermBuilder_GUI.Add("Text",,"Welcome to the term builder. `n")

    TermBuilder_GUI.Add("Text","section","Inst Num:")
    instNum_tbGUI := TermBuilder_GUI.Add("Edit","w60 yp Number Limit 2","1")
    instNum_tbGUI.OnEvent("Change",AutoTabGUI.Bind(,,2))

    TermBuilder_GUI.Add("Text","xs yp+32 w49","FICE:")
    FICE_tbgui := TermBuilder_GUI.Add("Edit","w80 yp Uppercase Limit6")
    FICE_tbgui.OnEvent("Change",AutoTabGUI.Bind(,,6))

    TermBuilder_GUI.Add("Text","section x+36 ys","Start row:")
    start_tbgui := TermBuilder_GUI.Add("Edit","w60 yp Number Limit 2","3")
    start_tbgui.OnEvent("Change",AutoTabGUI.Bind(,,2))

    TermBuilder_GUI.Add("Text","xs yp+32","End row:")
    end_tbgui := TermBuilder_GUI.Add("Edit","w62 yp Number Limit 2")
    end_tbgui.OnEvent("Change",AutoTabGUI.Bind(,,2))

    tbguiCLOSE := TermBuilder_GUI.Add("Button","NoTab","Home",)
    tbguiCLOSE.OnEvent("Click", tbToMain)
    
    tbguiGO := TermBuilder_GUI.Add("Button","yp","GO")
    tbguiGO.OnEvent("Click",tbGObtn.Bind(instNum_tbGUI,FICE_tbgui,start_tbgui,end_tbgui))

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& The Transcript Review GUI &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; SHATAEQ
TranscriptReview_GUI := Gui("AlwaysOnTop", "DTE  |  Transcript Review")
    TranscriptReview_GUI.SetFont("S" . FONT_SIZE, FONT_NAME)
    TranscriptReview_GUI.Add("Text",,"Welcome to the transcript reviewer. `n")

    TranscriptReview_GUI.Add("Text","section","Start row:")
    start_trgui := TranscriptReview_GUI.Add("Edit","w60 yp Number Limit 2","3")
    start_trgui.OnEvent("Change",AutoTabGUI.Bind(,,2))

    TranscriptReview_GUI.Add("Text","w68 xs yp+32","End row:")
    end_trgui := TranscriptReview_GUI.Add("Edit","w60 yp Number Limit 2")
    end_trgui.OnEvent("Change",AutoTabGUI.Bind(,,2))

    trguiCLOSE := TranscriptReview_GUI.Add("Button","NoTab","Home",)
    trguiCLOSE.OnEvent("Click", trToMain)
    
    trguiGO := TranscriptReview_GUI.Add("Button","yp","GO")
    trguiGO.OnEvent("Click",trGObtn.Bind(,,start_trgui,end_trgui))

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& The Articulator GUI &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; SHATATR
Articulator_GUI := Gui("AlwaysOnTop", "DTE  |  Articulator")
    Articulator_GUI.SetFont("S" . FONT_SIZE, FONT_NAME)
    Articulator_GUI.Add("Text",,"Welcome to the articulator tool. `n")

    Articulator_GUI.Add("Text","section","Start row:")
    start_artgui := Articulator_GUI.Add("Edit","w60 yp Number Limit 2","3")
    start_artgui.OnEvent("Change",AutoTabGUI.Bind(,,2))

    Articulator_GUI.Add("Text","w68 xs yp+32","End row:")
    end_artgui := Articulator_GUI.Add("Edit","w60 yp Number Limit 2")
    end_artgui.OnEvent("Change",AutoTabGUI.Bind(,,2))

    articCLOSE := Articulator_GUI.Add("Button","NoTab","Home",)
    articCLOSE.OnEvent("Click", articToMain)
    
    articGO := Articulator_GUI.Add("Button","yp","GO")
    articGO.OnEvent("Click",articGObtn.Bind(,,start_artgui,end_artgui))

; Pops up during Articulator to prompt the user about attributes
Attributes_GUI := Gui("AlwaysOnTop", "Add Attributes")
    Attributes_GUI.SetFont("S" . FONT_SIZE,FONT_NAME)
    Attributes_GUI.Add("Text",,"Select Equivalent Course Attributes. `n")

    ; This will be the map of attributes we may assign to an equivalent course
    ; Each attribute is represented by a Key in the Map
    ; If the value of a key is -1, then we do not add the attribute
    ; If the value of a key is +1, then we do add the attribute
    ; The callback function simply toggles by multiplying values by -1 OnChange
    MapOfAttr := Map()
    MapOfAttr["AL"] := -1
    MapOfAttr["SOSC"] := -1
    MapOfAttr["SCI"] := -1
    MapOfAttr["LSCI"] := -1
    MapOfAttr["ATSC"] := -1
    MapOfAttr["EGCD"] := -1
    MapOfAttr["HR"] := -1
    MapOfAttr["HE"] := -1
    MapOfAttr["PE"] := -1
    MapOfAttr["DEV"] := -1
    MapOfAttr["PT"] := -1
    MapOfAttr["WRIT"] := -1
    MapOfAttr["COMM"] := -1
    MapOfAttr["CCN"] := -1

    AttrALbtn := Attributes_GUI.Add("Checkbox",,"Arts and Letters")
    AttrALbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"AL"))

    AttrSOSCbtn := Attributes_GUI.Add("Checkbox",,"Social Sciences")
    AttrSOSCbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"SOSC"))

    AttrSCIbtn := Attributes_GUI.Add("Checkbox",,"Non-Lab Science")
    AttrSCIbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"SCI"))

    AttrLSCIbtn := Attributes_GUI.Add("Checkbox",,"Laboratory Science")
    AttrLSCIbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"LSCI"))

    AttrATSCbtn := Attributes_GUI.Add("Checkbox",,"Adv. Tech Science")
    AttrATSCbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"ATSC"))

    AttrEGCDbtn := Attributes_GUI.Add("Checkbox",,"Cultural Diversity")
    AttrEGCDbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"EGCD"))

    AttrHRbtn := Attributes_GUI.Add("Checkbox",,"Human Relations")
    AttrHRbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"HR"))

    AttrHEbtn := Attributes_GUI.Add("Checkbox",,"Health")
    AttrHEbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"HE"))

    AttrPEbtn := Attributes_GUI.Add("Checkbox",,"Physical Education")
    AttrPEbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"PE"))

    AttrPTbtn := Attributes_GUI.Add("Checkbox",,"Career-Technical")
    AttrPTbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"PT"))

    AttrDEVBbtn := Attributes_GUI.Add("Checkbox",,"Developmental")
    AttrDEVBbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"DEV"))

    AttrCCNbtn := Attributes_GUI.Add("Checkbox",,"Common Course No.")
    AttrCCNbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"CCN"))

    AttrHRbtn := Attributes_GUI.Add("Checkbox",,"AAOT Writing")
    AttrHRbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"WRIT"))

    AttrHRbtn := Attributes_GUI.Add("Checkbox",,"AAOT Communications")
    AttrHRbtn.OnEvent("Click",articPusher.Bind(,,MapOfAttr,"COMM"))


    addAttrBtn := Attributes_GUI.Add("Button",,"ADD")
    addAttrBtn.OnEvent("Click",articAddAttr.Bind(,,MapOfAttr))

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& The Program Change GUI &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
ProgramChange_GUI := Gui("AlwaysOnTop", "DTE  |  Program Changes")
    ProgramChange_GUI.SetFont("S" . FONT_SIZE,FONT_NAME)
    ProgramChange_GUI.Add("Text",,"Welcome to the program changer. `n")

    ; Make groupboxes for different high-level tasks
    ProgramChange_GUI.Add("GroupBox","section","Self-Assign Cases")
    CaseGrabber_BTN := ProgramChange_GUI.Add("Button","xs+5 ys+25","Assign case to self")

    ProgramChange_GUI.Add("GroupBox","section xs","Process Standard Cases")
    StandardSingleCase_BTN := ProgramChange_GUI.Add("Button","xs+5 ys+25","Process Standard Case (Single)")
    StandardMultiCase_BTN := ProgramChange_GUI.Add("Button","xp yp+40","Process Standard Case (Multiple)")
    StandardAddOnlyCase_BTN := ProgramChange_GUI.Add("Button","xp yp+40","Add Additional Program(s)")
    StandardUpdateOnlyCase_BTN := ProgramChange_GUI.Add("Button","xp yp+40","Update Area of Interest Code(s)")

    ProgramChange_GUI.Add("GroupBox","section xs","Process Urgent Cases")
    UrgentCase_BTN := ProgramChange_GUI.Add("Button","xs+5 ys+25","Process URGENT Case (Single)")


    posCLOSE := ProgramChange_GUI.Add("Button","NoTab","Home",)
    posCLOSE.OnEvent("Click", posToMain)
    
    posGO := ProgramChange_GUI.Add("Button","yp","GO")
    posGO.OnEvent("Click",posGObtn)

; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& The MAIN GUI &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
; &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
dteGui := Gui(,"DTE Helper Tool")
    dteGui.SetFont("S" . FONT_SIZE, FONT_NAME)
    dteGui.MarginX := Layout["marginX"]
    dteGui.MarginY := Layout["marginY"]

    ; Welcome Text
    dteGui.Add("Text",,"Welcome, what would you like to do? `n`n`n")
        
    ; --- GROUPBOX --- Transcript review processes
    dteGui.Add(
        "GroupBox",
        "section r4.4 w" . (Layout["fontSize"] * 20),
        "Transcript Review"
    )
    ; Articulate SHATATR
    articOPEN := dteGui.Add(
        "Button",
        "xs+" . Layout["smallGap"] . " ys+" . Layout["largeGap"],
        "Articulate courses"
    )
    articOPEN.OnEvent("Click",articOPENbtn)
    ; Term Builder SHATRNS
    tbOPEN := dteGui.Add(
        "Button",
        "xp",
        "Build transfer terms"
    )
    tbOPEN.OnEvent("Click",tbOPENbtn)
    ; Transcript Review SHATAEQ
    trOPEN := dteGui.Add(
        "Button",
        "xp",
        "Enter transfer course history"
    )
    trOPEN.OnEvent("Click",trOPENbtn)


    ; --- GROUPBOX --- Program change processes
    dteGui.Add(
        "GroupBox",
        "section r2 w" . (Layout["fontSize"] * 14) . " ys",
        "Program Changes"
    )
    ; Program Changes
    posOPEN := dteGui.Add(
        "Button", 
        "xs+" . Layout["smallGap"] . " ys+" . Layout["largeGap"],
        "Process program`nchange requests"
    )
    posOPEN.OnEvent("Click",posOPENbtn)

    ; --- GROUPBOX --- Diploma processes 
    dteGui.Add(
        "GroupBox",
        "section r2 w" . (Layout["fontSize"] * 12) . " ys",
        "Diplomas"
    )
    ; Diploma Dates SHADIPL
    diplOPEN := dteGui.Add(
        "Button",
        "xs+" . Layout["smallGap"] . " ys+" . Layout["largeGap"],
        "Enter outgoing`ndiploma dates"
    )
    diplOPEN.OnEvent("Click",diplOPENbtn)

    UsernameTagDTE := dteGui.Add("Text","section x420 y10","User: " . IniRead(GlobalConfigFile,"Local","User"))
    TermTagDTE := dteGui.Add("Text","section xs","Term: " . IniRead(GlobalConfigFile,"TermContext","CurrentTerm"))



; **********************************************************************************************************************
; ************************************************** HOT KEY SECTION ***************************************************
; **********************************************************************************************************************
; MAIN Hotkeys Ctrl + Alt + F__ for key features

; --- View Global Configs ---
^!F1::{
    ViewConfig()
}
; --- Access GUI to update Global Configs ---
^!F2::{
    UpdateConfigs()
}
; ^!F3::{
; }

; POS Change Hot Keys

^!F4::{
    global GlobalConfigFile
    LastName := StrUpper(SubStr(IniRead(GlobalConfigFile,"Local","User"),1,1)) . SubStr(IniRead(GlobalConfigFile,"Local","User"),2,StrLen(IniRead(GlobalConfigFile,"Local","User")) - 2)
    Press("{Tab}",2,50,200)

    Press(LastName,1,200,750)

    Press("{Tab}",10,0,200)

    Press("{Enter}",1,100,500)

    Press("{Tab}",6,0,200)

    Press("{Enter}",1,50,1000)
}

^!F5::{
    SendTermSnippet("Replaced ____ with ____ as primary program of study in {NextTerm} with same catalog term")
    return
}
^!F6::{
    SendTermSnippet("Replaced ____ with ____ as primary program of study in {NextTerm} with same catalog term; added ____ as additional program(s) of study")
    return
}
^!F7::{
    SendTermSnippet("Added ____  as additional program of study in {NextTerm} with same catalog term")
    return
}
^!F8::{
    SendTermSnippet("Updated AAOT area of interest Code from ____ to ____ in {NextTerm}")
    return
}
^!F9::{
    SendTermSnippet("Replaced ____ with ____ as primary program of study in {CurrentTerm} with same catalog term <URGENT REQUEST>")
    return
}

^!F10::{
    ProgramChange_GUI.Show("AutoSize Center")
}
; ^!F11::{
; }

; --- Reopen main gui after it has been closed
^!F12::{
    dteGui.Show("AutoSize Center")
}

; **********************************************************************************************************************
; *************************************************** DEBUG SECTION ****************************************************
; **********************************************************************************************************************

; **** Go Time ****
; UpdateGlobalStrings()
dteGui.Show("AutoSize Center")

; Alt + P pauses the application
!p::{
    Pause(-1)
    return
}
; Alt + C reloads ("crashes") the application
!c::{
    Reload ; !r is reserved in Banner for retrieving docs from BDMS
}

; Alt + Escape cancels the application
!Esc::{
    ExitApp
}