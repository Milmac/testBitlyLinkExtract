Imports System
Imports System.IO
Imports System.Net
Imports System.Text

Module Program
    Sub Main(args As String())

        'Dim bitlyUrl As String = "http://bit.ly/2UrBlXA"
        'Console.WriteLine(GetBitlyDestination(bitlyUrl))


        'Console.WriteLine(vbCrLf + "Is the destination of" + vbCrLf)
        'Console.WriteLine(bitlyUrl)

        'acts as a single row in the database table
        Dim htmlToParse As String = IO.File.ReadAllText("C:\Users\mille\source\repos\testBitlyLinkExtract\testBitlyLinkExtract\parseTestPage.html")

        'Console.WriteLine(ParseBitly(htmlToParse))
        'Console.WriteLine(ParseUtm(htmlToParse))
        Console.WriteLine(ParseUtm(ParseBitly(htmlToParse)))
        'ParseUtm(htmlToParse)

    End Sub

    Private Function ParseBitly(ByVal _html As String) As String

        'prepare html document for HtmlAgility 
        Dim document As New HtmlAgilityPack.HtmlDocument()
        document.OptionWriteEmptyNodes = True
        document.OptionFixNestedTags = True
        document.LoadHtml(_html)

        'bitly conversion result to return
        Dim result As String = _html

        'for each href link in the html document
        For Each link As HtmlAgilityPack.HtmlNode In document.DocumentNode.SelectNodes("//a[@href]")
            'Console.WriteLine(GetBitlyDestination(link.Attributes("href").Value))

            'replace bitly links with destination
            result = result.Replace(link.Attributes("href").Value, GetBitlyDestination(link.Attributes("href").Value))
            'get the inside of href with link.Attributes("href").Value
        Next

        Return result
    End Function

    Private Function ParseUtm(ByVal _html As String) As String

        'prepare html document for HtmlAgility 
        Dim document As New HtmlAgilityPack.HtmlDocument()
        document.OptionWriteEmptyNodes = True
        document.OptionFixNestedTags = True
        document.LoadHtml(_html)

        'bitly conversion result to return
        Dim result As String = _html

        Console.WriteLine("Old Links")
        'for each href link in the html document
        For Each link As HtmlAgilityPack.HtmlNode In document.DocumentNode.SelectNodes("//a[@href]")
            Console.WriteLine(link.Attributes("href").Value)
        Next

        Console.WriteLine("New Links")
        'for each href link in the html document
        For Each link As HtmlAgilityPack.HtmlNode In document.DocumentNode.SelectNodes("//a[@href]")
            'Console.WriteLine(link.Attributes("href").Value)

            'replace bitly links with destination
            result = result.Replace(link.Attributes("href").Value, RemoveUtm(link.Attributes("href").Value))
            'get the inside of href with link.Attributes("href").Value
        Next

        Return result
    End Function

    Private Function GetBitlyDestination(ByVal _bitlyUrl As String) As String

        'verify that it's a bitly link before

        'Make a Head request
        Dim request As WebRequest = WebRequest.Create(_bitlyUrl)
        request.Method = WebRequestMethods.Http.Head

        'Get response url from our bitly request
        Dim response As WebResponse = request.GetResponse()
        Return response.ResponseUri.ToString()

        'New Ippea.core.uri(response.ResponseUri).linkuri.ToString()
        'add a check to remove non domain links
    End Function


    Private Function RemoveUtm(ByVal _utmUrl As String) As String

        'add baseurl to internal links to fix Uri error
        Dim baseUrl As String = "https://www.inyopools.com"

        'clean utmlUrl by removing unnecessary "amp;" as this messes with the ParseQueryString Method
        Dim clean_utmUrl = _utmUrl.Replace("amp;", "")



        'create a new uri using the input url
        Dim uri As Uri
        Dim isAbsoluteResult As Boolean = IsAbsoluteUrl(clean_utmUrl)

        If isAbsoluteResult Then
            Console.WriteLine("Absolute Url Detected!!")
            uri = New Uri(clean_utmUrl)
        Else
            Console.WriteLine("Relative Url Detected!")
            'add inyopools.com to fix parse query string method since it only works for absolute urls
            uri = New Uri(baseUrl + clean_utmUrl)
        End If



        'parse the queries in the uri
        Dim newQueryString As Specialized.NameValueCollection = System.Web.HttpUtility.ParseQueryString(uri.Query)

        'utm querys to be removed
        newQueryString.Remove("utm_source")
        newQueryString.Remove("utm_medium")
        newQueryString.Remove("utm_content")
        newQueryString.Remove("utm_campaign")

        'store the left side of the _utmUrl without any queries. Also remove inyopools base url as we want it to be a relative url
        Dim pagePathWithoutQueryString As String = Replace(uri.GetLeftPart(UriPartial.Path), "https://www.inyopools.com", "")

        'if there are more than 0 queries that still exist
        If newQueryString.Count > 0 Then
            'Add ? between left and right side of new url
            Console.WriteLine(String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString))
            Return String.Format("{0}?{1}", pagePathWithoutQueryString, newQueryString)
        Else 'return just the left side since there are no queries to add
            Console.WriteLine(pagePathWithoutQueryString)
            Return pagePathWithoutQueryString
        End If

    End Function

    'returns true if the url is an absolute url and false if it's a relative url
    Private Function IsAbsoluteUrl(ByVal url As String) As Boolean
        Dim result As Uri

        Return Uri.TryCreate(url, UriKind.Absolute, result)
    End Function
End Module
