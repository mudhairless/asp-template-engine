<%
'Template Engine

class TemplateEngine

    private sub Class_Initialize
        set m_items = Server.CreateObject("Scripting.Dictionary")
        set m_commands = Server.CreateObject("Scripting.Dictionary")

        m_commands.add "IF", "\{\{IF\(.*\)\}\}"
        m_commands.add "ELSE", "\{\{ELSE\}\}"
        m_commands.add "ENDIF", "\{\{ENDIF\}\}"
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

    private function m_getFileContentsAsArray( fnHTMLfile )

        dim fnResultArray(), fnLineCount, fnfso, fnFile, fnLine
        fnLineCount = 0

        set fnfso = Server.CreateObject("Scripting.FileSystemObject")

            if(fnfso.FileExists(Server.MapPath(m_templ_dir & fnHTMLfile))) then

                set fnFile = fnfso.OpenTextFile(Server.MapPath(m_templ_dir & fnHTMLfile))

                    while not fnFile.EOF

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

    private function m_getFileContents( fnHTMLfile )

        dim fnResult, fnfso, fnFile

        set fnfso = Server.CreateObject("Scripting.FileSystemObject")

            if(fnfso.FileExists(Server.MapPath(m_templ_dir & fnHTMLfile))) then

                set fnFile = fnfso.OpenTextFile(Server.MapPath(m_templ_dir & fnHTMLfile))

                    fnResult = fnFile.ReadAll()

                    if(fnResult = "") then
                        fnResult = "Error reading template file: " & fnHTMLfile
                    end if

                set fnFile = nothing

            else

                fnResult = false

            end if

        set fnfso = nothing

        m_getFile = fnResult

    end function

    public function parse( fnHTMLfile )

        dim fnFileContents, fnResult, fnFileLine, fnregex, fnCommand, fnIfLevel
        fnIfLevel = 0

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
                                case "ENDIF"
                                    fnIfLevel = fnIfLevel - 1
                            end select

                        end if

                    set fnregex = nothing
                next

            next

        else
            fnResult = "Could not load template file: " & m_templ_dir & fnHTMLfile
        end if

        if(fnIfLevel > 0) then
            fnResult = "Expected {{ENDIF}} Source: " & m_templ_dir & fnHTMLfile
        elseif(fnIfLevel < 0) then
            fnResult = "Unexpected {{ENDIF}} Source: " & m_templ_dir & fnHTMLfile
        end if

        parse = fnResult

    end function

    public function apply( fnHTMLfile )

        dim fnResult, fnItem, fnIV, fnPattern, i, fnArrRes

        fnResult = m_getFileContents(fnHTMLfile)

        if (fnResult <> false) then

            for each fnItem in m_items

                fnIV = m_items(fnItem)

                if(isArray(fnIV)) then

                    fnPattern = fnIV(0)
                    fnArrRes = ""

                    for i = 1 to ubound(fnIV)

                        fnArrRes = Replace(fnArrRes,"{{ITEM}}",fnIV(i))
                        fnArrRes = Replace(fnArrRes,"{{INDEX}}",i-1)

                    next

                    fnResult = Replace(fnResult,"{{" & fnItem & "}}",fnArrRes)

                else

                    fnResult = Replace(fnResult,"{{" & fnItem & "}}",fnIV)

                end if

            next

        else

            fnResult = "Could not load template file: " & m_templ_dir & fnHTMLfile

        end if

        apply = fnResult

    end function

    private m_items
    private m_templ_dir
    private m_commands


end class

%>
