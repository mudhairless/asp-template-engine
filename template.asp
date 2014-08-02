<%
'Template Engine

class TemplateEngine

    private sub Class_Initialize
        set m_items = Server.CreateObject("Scripting.Dictionary")
        set m_commands = Server.CreateObject("Scripting.Dictionary")

        m_commands.add "IF", "\{\{IF\(.*\)\}\}"
        m_commands.add "ELSE", "\{\{ELSE\}\}"
        m_commands.add "ENDIF", "\{\{ENDIF\}\}"
        m_commands.add "INCLUDE", "\{\{\INCLUDE\[.*\]\}\}"

        m_warn_on_unused = false
    end sub

    private sub Class_Deinitialize
        set m_items = nothing
        set m_commands = nothing
    end sub

    public property get Items
        Items = m_items
    end property

    public sub add( fnKey, fnVal )
        m_items.add fnKey, fnVal
    end sub

    public property Let TemplateDirectory( fnDir )
        m_templ_dir = fnDir
    end property

    public property get TemplateDirectory
        TemplateDirectory = m_templ_dir
    end property

    public property Let WarnOnUnusedTags( fBool )
        m_warn_on_unused = fbool
    end property

    public property get WarnOnUnusedTags
        WarnOnUnusedTags = m_warn_on_unused
    end property

    private function m_getFileContentsAsArray( fnHTMLfile )

        dim fnResultArray(), fnLineCount, fnfso, fnFile, fnLine
        fnLineCount = 0

        set fnfso = Server.CreateObject("Scripting.FileSystemObject")

            if(fnfso.FileExists(Server.MapPath(m_templ_dir & fnHTMLfile))) then

                set fnFile = fnfso.OpenTextFile(Server.MapPath(m_templ_dir & fnHTMLfile))

                    while not fnFile.AtEndOfStream

                        fnLine = fnFile.ReadLine()
                        fnLineCount = fnLineCount + 1

                        redim preserve fnResultArray(fnLineCount)

                        fnResultArray(fnLineCount-1) = fnLine

                    wend

                set fnFile = nothing

            end if

        set fnfso = nothing

        if(fnLineCount = 0) then
            redim fnResultArray(1)
            fnResultArray(0) = false
        end if

        m_getFileContentsAsArray = fnResultArray

    end function

    private function m_getFileContents( fnHTMLfilef )

        dim fnResultf, fnfsof, fnFilef

        set fnfsof = Server.CreateObject("Scripting.FileSystemObject")

            if(fnfsof.FileExists(Server.MapPath(m_templ_dir & fnHTMLfilef))) then

                set fnFilef = fnfsof.OpenTextFile(Server.MapPath(m_templ_dir & fnHTMLfilef))

                    fnResultf = fnFilef.ReadAll()

                    if(fnResultf = "") then
                        fnResultf = "Error reading template file: " & fnHTMLfilef
                    end if

                set fnFilef = nothing

            else

                fnResultf = false

            end if

        set fnfsof = nothing

        m_getFileContents = fnResultf

    end function

    private function m_applyToString( mFnString )

        dim mFnResult, fnItem, fnIV, fnPattern, i, fnArrRes
        mFnResult = mFnString

        for each fnItem in m_items

            fnIV = m_items(fnItem)

            if(isArray(fnIV)) then

                fnPattern = fnIV(0)
                fnArrRes = ""

                for i = 1 to ubound(fnIV)

                    fnArrRes = Replace(fnArrRes,"{{ITEM}}",fnIV(i))
                    fnArrRes = Replace(fnArrRes,"{{INDEX}}",i-1)

                next

                mFnResult = Replace(mFnResult,"{{" & fnItem & "}}",fnArrRes)

            else

                mFnResult = Replace(mFnResult,"{{" & fnItem & "}}",fnIV)

            end if

        next

        m_applyToString = mFnResult

    end function

    private function m_checkForUnusedTags( mFnString )
        dim mFnResult, mFnregex
        set mFnregex = new RegExp

            mFnregex.Pattern = "\{\{.*\}\}"
            mFnregex.Global = true

            if(mFnregex.Test(mFnString)) then
                set mFnMatches = mFnregex.Execute(mFnString)
                for each mfnmatch in mFnMatches
                    mFnResult = mFnResult & "Tag " & mfnmatch.Value & " still exists in output." & vbCRLF
                next
            else
                mFnResult = mFnString
            end if

        set mFnregex = nothing

        m_checkForUnusedTags = mFnResult

    end function

    public function parse( fnHTMLfile )

        dim fnFileContents, fnResult, fnFileLine, fnregex, fnCommand, fnIfLevel, fnDoOutput(), fnTestValue
        fnIfLevel = 0
        redim fnDoOutput(1)
        fnDoOutput(0) = true

        fnFileContents = m_getFileContentsAsArray(fnHTMLfile)

        if(fnFileContents(0) <> false) then

            for each fnFileLine in fnFileContents

                for each fnCommand in m_commands
                    set fnregex = new RegExp

                        fnregex.Pattern = m_commands(fnCommand)
                        fnregex.IgnoreCase = true
                        fnregex.Global = true

                        if(fnregex.Test(fnFileLine)) then

                            select case fnCommand
                                case "IF":
                                    fnIfLevel = fnIfLevel + 1
                                    redim preserve fnDoOutput(fnIfLevel)
                                    fnDoOutput(fnIfLevel) = true
                                    set fnMatches = fnregex.Execute(fnFileLine)
                                    for each fnMatch in fnMatches
                                        fnTestValue = fnMatch.Value
                                        fnTestValue = Mid(fnTestValue,6,len(fnTestValue)-8)
                                        if(instr(fnTestValue,"=") > 0) then
                                            fnTestVals = Split(fnTestValue,"=")
                                            if(m_items(fnTestVals(0)) <> fnTestVals(1)) then
                                                fnDoOutput(fnIfLevel) = false
                                            end if
                                        else
                                            if(NOT m_items.Exists(fnTestValue)) then
                                                fnDoOutput(fnIfLevel) = false
                                            end if
                                        end if
                                    next
                                    set fnMatches = nothing
                                    fnFileLine = fnregex.Replace(fnFileLine,"")
                                case "ENDIF"
                                    fnIfLevel = fnIfLevel - 1
                                    fnFileLine = fnregex.Replace(fnFileLine,"")
                                case "ELSE"
                                    if(fnIfLevel = 0) then
                                        fnResult = "Unexpected {{ELSE}}"
                                        exit for
                                    end if
                                    if(fnDoOutput(fnIfLevel)) then
                                        fnDoOutput(fnIfLevel) = false
                                    else
                                        fnDoOutput(fnIfLevel) = true
                                    end if
                                    fnFileLine = fnregex.Replace(fnFileLine,"")

                                case "INCLUDE"
                                    if(fnDoOutput(fnIfLevel)) then
                                    set fnMatches = fnregex.Execute(fnFileLine)
                                        for each fnMatch in fnMatches
                                            fnTestValue = fnMatch.Value
                                            fnTestValue = Mid(fnTestValue,11,len(fnTestValue)-13)
                                            fnFileLine = fnregex.Replace(fnFileLine,parse(fnTestValue))
                                        next
                                        set fnMatches = nothing

                                    end if


                            end select

                        end if

                    set fnregex = nothing
                next

                if(fnDoOutput(fnIfLevel)) then
                    fnResult = fnResult & fnFileLine
                end if

            next

        else
            fnResult = "Could not load template file: " & m_templ_dir & fnHTMLfile
        end if

        if(fnIfLevel > 0) then
            fnResult = "Expected {{ENDIF}} Source: " & m_templ_dir & fnHTMLfile
        elseif(fnIfLevel < 0) then
            fnResult = "Unexpected {{ENDIF}} Source: " & m_templ_dir & fnHTMLfile
        end if

        fnResult = m_applyToString(fnResult)

        if(m_warn_on_unused) then

            fnResult = m_checkForUnusedTags(fnResult)

        end if

        parse = fnResult

    end function

    public function apply( fnHTMLfile )

        dim fnResult

        fnResult = m_getFileContents(fnHTMLfile)

        if (not isEmpty(fnResult) and fnResult <> false) then

            fnResult = m_applyToString(fnResult)

        else

            fnResult = "Could not load template file: " & Server.MapPath(m_templ_dir & fnHTMLfile) & "<br/>"

        end if

        if(m_warn_on_unused) then

            fnResult = m_checkForUnusedTags(fnResult)

        end if

        apply = fnResult

    end function

    private m_items
    private m_templ_dir
    private m_commands
    private m_warn_on_unused


end class

%>
