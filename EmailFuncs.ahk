#NoEnv
#SingleInstance Force

; Fetches current selected email within Outlook
GetCurrentEmail() {
    return ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
}

; Fetches the Full Name of the sender of the currently selected email
GetSender() {
    return GetCurrentEmail().SenderName
}

; Converts a Full Name into a Standard Name by removing the Middle Initial
; If none passed in, defaults to sender of Current Email
GetStandardName(name:="") {
    if(name = "") {
        name := GetSender()
    }
    return RegExReplace(name,"\w\. ")
}

; Extracts the first name from a Standard Name
; If none, defaults to sender of Current email
GetFirstName(name:="") {
    if(name = "") {
        name := GetStandardName()
    }
    return RegExReplace(name,"\s.*")
}

; Trims out the contents of an email after the name provided
; If no sender provided, passes empty string internally
; If no email provided, uses current selected email in outlook
GetEmailBody(email:="",name:="") {
    whitespace := " `t`n`r"
    regexstr := "s)(?=From:|" . GetFirstName(name) . "|" . GetStandardName(name) . "|"
    linecleaner := "\s{2,}"
    
    if(email = "") {
        email := GetCurrentEmail().Body
    }

    if(name = "") {
        regexstr := regexstr . GetSender() . ").*"
    } else {
        regexstr := regexstr . name . ").*"
    }
    return Trim(RegExReplace(RegExReplace(email, regexstr), linecleaner, "`n`n"), whitespace)
}

; Sets the Category of the email to mark that it was handled by me, and marks as read
SetAsHandled() {
    email := GetCurrentEmail()
    email.Categories := "ST"
    email.UnRead := False
    return
}

; Returns the domain of an email address
GetEmailDomain(address:="") {
    if(address = "") {
        address := GetCurrentEmail().SenderEmailAddress
    }
    RegExMatch(address, "(?<=@).*(?=\.)", account)
    return account
}

; Returns a random boilerplate email greeting
GenerateGreeting() {
    greetings := ["Hi","Hey"]
    Random, idx, 1, greetings.Length()
    return greetings[idx]
}
