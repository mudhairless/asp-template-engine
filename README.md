CATE: Classic ASP Template Engine
=================================

CATE is a simple but powerful template engine for Classic ASP with an easy to learn syntax similar to other template engines.

CATE Simple Template Syntax Example
===================================

    {{if(PAGE_TITLE)}}
       {{INCLUDE[header.tpl]}}
    {{endif}}
    {if(test=true)}}
       {{PAGE_CONTENT}}
    {{else}}
       No content to display.
    {{endif}}
    {{if(PAGE_TITLE)}}
       {{INCLUDE[footer.tpl]}}
    {{endif}}

How To Use CATE from ASP
========================
    <!-- #include virtual="/path/to/template.asp" -->
    <%
        set TENGINE = new TemplateEngine
        TENGINE.TemplateDirectory = "/path/to/templates/"
        TENGINE.add "PAGE_TITLE", "Test"
        TENGINE.add "PAGE_CONTENT", "Testing"
        TENGINE.add "test", "true"
        outcontent = TENGINE.parse("testing.tpl")
        Response.Write outcontent
        set TENGINE = nothing
    %>

Methods
=======

  * add( variable_name, value )

    > Adds the variable to the list of items to be replaced in templates.
      Maps to {{variable_name}} in the template. Returns true if the
      variable was newly added or false if the variable already exists.

  * apply( template_file )

    > Performs a simple substitution of all named variables in the
      specified template file. Does not handle commands like if or include
      but is much faster. Returns a string containing the template file
      with variables replaced or an error message.

  * parse( template_file )

    > Similar to apply but also parses the template file for commands.
      Returns a string containing the template file with commands parsed
      and variables replaced or an error message.

  * replace( variable_name, value )

    > Replaces the named variables value with the specified value.
      Returns true if the variable was replaced or false if it does not exist.

Properties
==========

  * TemplateDirectory (read/write)

    > Sets the base directory to look for template files in.
      When looking for templates the engine itself calls Server.MapPath on the
      path so all paths should be virtual, i.e. relative to the wwwroot
      directory. Default = ""

  * WarnOnUnusedTags (read/write)

    > This option controls whether the engine will output an error message
      if unparsed tags are left in its output. Set to true to enable or
      false to disable. Default = false

  * Items (read-only)

    > Returns the underlying Dictionary collection storing the variables and their values.

Commands
========

* `{{IF(expression)}}`

    > Outputs the content following the IF statement until an ELSE or
      ENDIF statement is reached only if the expression is true. The expression
      takes the following forms:

    * `variable_name=value`

          > The named variable's value is tested against the string value on the
            right side of the equal sign (=) and the expression is considered
            true if they match.

    * `variable_name`

          > This expression form is considered true if the named variable
            exists in the Items collection.

* `{{ELSE}}`

    > Outputs the content following the ELSE statement until an ENDIF
      statement is reached if the preceding IF statement expression
      evaluates to false.

* `{{ENDIF}}`

    > Ends an IF statement.

* `{{INCLUDE[template_file]}}`

    > Parses the contents of the specified file and places them at the point
      of the INCLUDE statement.
