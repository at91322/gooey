#Requires AutoHotkey v2.0
#SingleInstance force

; --- Global Configurations ---
; For now these are a mirror of the configuration setup I'm using for the application
; that needs this GUI project

GlobalConfigFile := A_ScriptDir "\config.ini"
global FONT_SIZE := IniRead(GlobalConfigFile,"Appearance","FontSize")
global FONT_NAME := IniRead(GlobalConfigFile,"Appearance","Font")
global THEME := IniRead(GlobalConfigFile,"Appearance","Theme")

; --- Class Section ---

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

; --- Function Section ---
ExampleUsage(){
    ; Create a new GUI builder
    builder := GUIBuilder("Dynamic GUI Example", "Resize")

    ; Create a GroupBox for user information
    userGroup := builder.CreateGroupBox("userInfo", "User Information", 10, 10, 350, 150)

    ; Add controls to the GroupBox
    builder.AddTextEditPair("userInfo", "firstName", "First Name:", 200, "John")
    builder.AddTextEditPair("userInfo", "lastName", "Last Name:", 200, "Doe")
    builder.AddTextDropDownListPair("userInfo", "country", "Country:", ["USA", "Canada", "UK", "Germany"], 200, 1)
    builder.AddTextCheckBoxPair("userInfo", "newsletter", "Options:", "Subscribe to newsletter", true, 200)

    ; Create another GroupBox
    prefsGroup := builder.CreateGroupBox("preferences", "Preferences", 10, , 350, 100)
    builder.AddTextDropDownListPair("preferences", "theme", "Theme:", ["Light", "Dark", "Auto"], 200, 2)
    builder.AddTextCheckBoxPair("preferences", "notifications", "Alerts:", "Enable notifications", false, 200)

    ; Add standard buttons
    buttons := builder.AddStandardButtons(["OK", "Cancel", "Apply"])
    
    ; Set up button events
    buttons["OK"].OnEvent("Click", OkClick.Bind(builder))
    buttons["Cancel"].OnEvent("Click", CancelClick.Bind(builder))
    
    ; Show the GUI
    builder.Show("w380 h340")
    
    return builder
}

UpdateConfigs(){
     ; Create a new GUI builder
    builder := GUIBuilder("DTE  |  Update Configurations", "Resize")

    ; Create a GroupBox for local configuration information
    localGroup := builder.CreateGroupBox("localInfo", "Local Settings", 10, 10, 350, 69)

    ; Add controls to the GroupBox
    builder.AddTextEditPair("localInfo", "username", "Username:", 100, IniRead(GlobalConfigFile,"Local","Username"))

    ; Create a GroupBox for term context configuration information
    termGroup := builder.CreateGroupBox("termInfo", "Term Settings", 10, , 350, 69)

    ; Add controls to the GroupBox
    builder.AddTextEditPair("termInfo", "currentTerm", "Current Term:", 62, IniRead(GlobalConfigFile,"TermContext","CurrentTerm"))

    ; Create a GroupBox for session configuration information
    sessionGroup := builder.CreateGroupBox("sessionInfo", "Session Settings", 10, , 350, 69)

    ; Add controls to the GroupBox
    builder.AddTextDropDownListPair("sessionInfo", "systemSpeed", "System Speed:", ["1 (Normal)","2","3","4","5 (Slow)"], 100, IniRead(GlobalConfigFile,"SessionContext","SystemSpeed"))

    ; Create a GroupBox for appearance configuration information
    appearanceGroup := builder.CreateGroupBox("appearanceInfo", "Appearance Settings", 10, , 350, 121)

    ; Add controls to the GroupBox
    builder.AddTextDropDownListPair("appearanceInfo", "fontName", "Font:", ["Lucia Sans"], 200, 1)
    builder.AddTextDropDownListPair("appearanceInfo", "fontSize", "Font Size:", ["8","10","12","14","16"], 200, 3)
    builder.AddTextDropDownListPair("appearanceInfo", "theme", "Theme:", ["Light", "Dark"], 200, 1)

    ; Add standard buttons
    buttons := builder.AddStandardButtons(["OK", "Cancel", "Apply"])

    ; Set up button events
    buttons["Apply"].OnEvent("Click", UpdateConfigApplyClick.Bind(builder))
    buttons["Cancel"].OnEvent("Click", CancelClick.Bind(builder))
    
    ; Show the GUI
    builder.Show("AutoSize Center")
    
    return builder
}
; Use IniWrite to update config.ini
UpdateConfigApplyClick(builderRef, *){
}

OkClick(builderRef, *) {
    ; Demonstrate getting values
    firstName := builderRef.GetValue("firstName")
    country := builderRef.GetValue("country")
    newsletter := builderRef.GetValue("newsletter")

    message := "First Name: " . firstName . "`nCountry: " . country . "`nNewsletter: " . newsletter
    MsgBox(message)
}

CancelClick(builderRef, *) {
    builderRef.gui.Hide()
}


; --- Hot Key Section ---

^!F1::{
    ExampleUsage()
}
^!F2::{
    UpdateConfigs()
}
; ^!F3::{
; }
; ^!F4::{
; }
; ^!F5::{
; }
; ^!F6::{
; }
; ^!F7::{
; }
; ^!F8::{
; }
; ^!F9::{
; }
; ^!F10::{
; }
; ^!F11::{
; }
; ^!F12::{
; }

!p::{
    Pause(-1)
}

!c::{
    Reload
}

!Esc::{
    ExitApp
}