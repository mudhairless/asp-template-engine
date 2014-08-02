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