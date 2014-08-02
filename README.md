CATE: Classic ASP Template Engine
=================================

CATE is a simple but powerful template engine for Classic ASP.
CATE has a easy to learn syntax similar to other template engines.

CATE Template Syntax Example
============================

    {{if(PAGE_TITLE)}}
       {{INCLUDE[header.tpl]}}
    {{endif}}
    {if(test=true)}}
       {{PAGE_CONTENT}}
    {{else}}
       No content to display.
    {{endif}}
    {{INCLUDE[footer.tpl]}}

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
        Response.Write outcontent & "<h3>" & timeout - timein & "</h3>"
        set TENGINE = nothing
    %>

