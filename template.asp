<!-- #include file="md5.asp" -->
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

        set m_hasher = new MD5
        m_enable_cache = false
        m_cache_filename = "{{VARHASH}}.{{VALHASH}}.{{FILE}}.cache"
        m_cache_dir = "/cache/"
    end sub

    private sub Class_Deinitialize
        set m_items = nothing
        set m_commands = nothing
        set m_hasher = nothing
    end sub

    public property Let EnableCache( mbool )
        m_enable_cache = mbool
        if(m_enable_cache = true) then
            call hash_all()
        end if
    end property

    public property get EnableCache
        EnableCache = m_enable_cache
    end property

    public property Let CacheDirectory( mcstring )
        m_cache_dir = mcstring
    end property

    public property get CacheDirectory
        CacheDirectory = m_cache_dir
    end property

    public property Let CacheFilename( mcstring )
        m_cache_filename = mcstring
    end property

    public property get CacheFilename
        CacheFilename = m_cache_filename
    end property

    public property get Items
        Items = m_items
    end property

    public function add( fnKey, fnVal )
        if(not m_items.Exists(fnKey)) then
            m_items.add fnKey, fnVal
            if(m_enable_cache) then
                call hash_new(fnKey,fnVal)
            end if
            add = true
        else
            add = false
        end if
    end function

    public function replaceValue( fnKey, fnVal )
        if(not m_items.Exists(fnKey)) then
            replaceValue = false
        else
            m_items(fnKey) = fnVal
            replaceValue = true
            m_input_hash_vars = ""
            m_input_hash_values = ""
            call hash_all()
        end if
    end function

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

    private sub hash_all
        dim fnKey, fnVal, fnVali
        for each fnKey in m_items
            m_input_hash_vars = m_hasher.Hash(m_input_hash_vars & fnKey)
            fnVal = m_items(fnKey)
            if(isArray(fnVal)) then
                for each fnVali in fnVal
                    m_input_hash_values = m_hasher.Hash(m_input_hash_values & fnVali)
                next
            else
                m_input_hash_values = m_hasher.Hash(m_input_hash_values & fnVal)
            end if
        next
    end sub

    private sub hash_new( mkey, mval )
        dim fnVali
        m_input_hash_vars = m_hasher.Hash(m_input_hash_vars & mkey)
        if(isArray(mval)) then
            for each fnVali in mval
                m_input_hash_values = m_hasher.Hash(m_input_hash_values & fnVali)
            next
        else
            m_input_hash_values = m_hasher.Hash(m_input_hash_values & mval)
        end if
    end sub

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

                    fnFile.Close

                set fnFile = nothing

            end if

        set fnfso = nothing

        if(fnLineCount = 0) then
            redim fnResultArray(1)
            fnResultArray(0) = false
        end if

        m_getFileContentsAsArray = fnResultArray

    end function

    private sub m_writeFile( fnFileToWrite, fnFileData )

        dim mfnfso, mfnfileh
        set mfnfso = Server.CreateObject("Scripting.FileSystemObject")

            set mfnfileh = mfnfso.CreateTextFile(Server.MapPath(fnFileToWrite),2,0)

                if(isArray(fnFileData)) then
                    for each mFileLine in fnFileData
                        mfnfileh.WriteLine(fnFileData)
                    next
                else
                    mfnfileh.Write(fnFileData)
                end if

                mfnfileh.Close

            set mfnfileh = nothing

        set mfnfso = nothing

    end sub

    private function m_getFileContents( fnHTMLfilef )

        dim fnResultf, fnfsof, fnFilef

        set fnfsof = Server.CreateObject("Scripting.FileSystemObject")

            if(fnfsof.FileExists(Server.MapPath(fnHTMLfilef))) then

                set fnFilef = fnfsof.OpenTextFile(Server.MapPath(fnHTMLfilef),1,true,0)

                    fnResultf = fnFilef.ReadAll()

                    if(fnResultf = "") then
                        fnResultf = "Error reading template file: " & fnHTMLfilef
                    end if

                    fnFilef.Close

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

        dim fnFileContents, fnResult, fnFileLine, fnregex, fnCommand, fnIfLevel, fnDoOutput(), fnTestValue,fnLineNo,fnSkip
        fnIfLevel = 0
        redim fnDoOutput(1)
        fnDoOutput(0) = true
        fnLineNo = 1
        fnSkip = false

        if(m_enable_cache) then

            dim fnCachedFilename
            fnCachedFilename = Replace(m_cache_filename,"{{VARHASH}}", m_input_hash_vars)
            fnCachedFilename = Replace(fnCachedFilename,"{{VALHASH}}", m_input_hash_values)
            fnCachedFilename = Replace(fnCachedFilename,"{{FILE}}", fnHTMLfile)

            fnFileContents = m_getFileContents(m_cache_dir & fnCachedFilename)

            if(fnFileContents <> false AND fnFileContents <> "") then
                fnResult = fnFileContents
                fnSkip = true
            end if

        end if

        if(not fnSkip) then

            fnFileContents = m_getFileContentsAsArray(m_templ_dir & fnHTMLfile)

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

                    fnLineNo = fnLineNo + 1

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

            if(m_enable_cache) then
                fnCachedFilename = Replace(m_cache_filename,"{{VARHASH}}", m_input_hash_vars)
                fnCachedFilename = Replace(fnCachedFilename,"{{VALHASH}}", m_input_hash_values)
                fnCachedFilename = Replace(fnCachedFilename,"{{FILE}}", fnHTMLfile)
                call m_writeFile(m_cache_dir & fnCachedFilename,fnResult)
            end if

        end if

        parse = fnResult

    end function

    public function apply( fnHTMLfile )

        dim fnResult, fnSkip, fnFileContents
        fnSkip = false

        if(m_enable_cache) then

            dim fnCachedFilename
            fnCachedFilename = Replace(m_cache_filename,"{{VARHASH}}", m_input_hash_vars)
            fnCachedFilename = Replace(fnCachedFilename,"{{VALHASH}}", m_input_hash_values)
            fnCachedFilename = Replace(fnCachedFilename,"{{FILE}}", fnHTMLfile)

            fnFileContents = m_getFileContents(m_cache_dir & fnCachedFilename)

            if(fnFileContents <> false AND fnFileContents <> "") then
                fnResult = fnFileContents
                fnSkip = true
            end if

        end if

        if(not fnSkip) then

            fnResult = m_getFileContents(m_templ_dir & fnHTMLfile)

            if (not isEmpty(fnResult) and fnResult <> false) then

                fnResult = m_applyToString(fnResult)

            else

                fnResult = "Could not load template file: " & Server.MapPath(m_templ_dir & fnHTMLfile) & "<br/>"

            end if

            if(m_warn_on_unused) then

                fnResult = m_checkForUnusedTags(fnResult)

            end if

            if(m_enable_cache) then
                fnCachedFilename = Replace(m_cache_filename,"{{VARHASH}}", m_input_hash_vars)
                fnCachedFilename = Replace(fnCachedFilename,"{{VALHASH}}", m_input_hash_values)
                fnCachedFilename = Replace(fnCachedFilename,"{{FILE}}", fnHTMLfile)
                call m_writeFile(m_cache_dir & fnCachedFilename,fnResult)
            end if

        end if

        apply = fnResult

    end function

    private m_items
    private m_templ_dir
    private m_commands
    private m_warn_on_unused
    private m_input_hash_vars
    private m_input_hash_values
    private m_enable_cache
    private m_cache_dir
    private m_hasher
    private m_cache_filename

end class

%>
