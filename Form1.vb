Imports System.Text.RegularExpressions

Public Class Form1

    'REGULAR EXPRESSIONS'
    '-------------------'

    '***NOTE: "n/a" => NO TEXT TO MATCH IN SOURCE CODE'

    '#1 
    'VARIABLE NAME: task  
    'REGEX CODE: (?>task\s*:\s*?)([0-9a-z,.:;@'""\/\-_!£$%^&*()\[\]~# ]+)(?><br\s*\/\s*>)?'
    'EXPECTED RESULTS: 1192,1192,n/a, n/a'
    'SAMPLE TEXT: [1] Task:1192<br />, [2] <br><hr><br><br>      Task:1192<br><br>'
    'TEST LINKS: 1: y, 2: y, 3: n/a , 4: n/a'


    '#2 
    'VARIABLE NAME: text 
    'REGEX CODE: (?>text\s*:\s*?)([0-9a-z,.\/?:;~#\[\]()!""£$%^&*|\-_ ]+)'
    'EXPECTED RESULT: Create task.'
    'SAMPLE TEXT: [1] Text:Create task.<br />, [2] Text:Create task.<br><br> '
    'TEST LINKS: 1: y, 2: y, 3: n/a, 4: n/a


    '#3 
    'VARIABLE NAME: time 
    'REGEX CODE: (?>(?>time|sendt)\s*:\s*)(.*?\w*\s*(?>19|20)\d{2}\s*)?(\d{1,2}\s*(?>:|;|,|.|-|_)\s*\d{1,2}\s*)(?><\s*br\s*\/>)?'
    'EXPECTED RESULT: 12,5, 12:16'
    'SAMPLE TEXT: [1] Time:12,5<br />, [2] <br><br>     Time:12,5<br><br><br', [3] n/a Sendt: 26. august 2015 12:16
    'TEST LINKS: 1: y, 2: y , 3: n/a, 4:  y


    '#4 
    'VARIABLE NAME: fromEmail 
    'REGEX CODE: (?:(?>from|fra)\s*:\s*)(?:[a-z0-9\/|<>,.?:;'~#!""£$%^&*()\-_\+=\[\]\+ ]+)?(?:(?:\[mailto\s*:\s*)?(?><\s*a\s*href=""|\[)?mailto\s*:\s*)(\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b)(?><\s*mailto\s*:\s*|"">|\])?'
    'EXPECTED RESULT: [1] rr@webia.dk, [2] pka@nyfors.dk', [3] n/a, [4] rr@webia.dk
    'SAMPLE TEXT: [1] From: Robert Radut [mailto:rr@webia.dk <mailto:rr@webia.dk> ]<br>, 
    '[2] <br>From: <b class="gmail_sendername">Per Kaptain</b> <span dir="ltr">&lt;<a href="mailto:pka@nyfors.dk">pka@nyfors.dk</a>', [3] n/a, 
    '[4] Fra:</span></b><span style="font-size:10.0pt;font-family:&quot;Tahoma&quot;,&quot;sans-serif&quot;"> Robert Radut - Webia [mailto:<a href="mailto:rr@webia.dk" target="_blank">rr@webia.dk</a>]
    'TEST LINKS: 1: y, 2: y, 3: n/a , 4: y '


    '#5 
    'VARIABLE NAME: fromPerson 
    'REGEX CODE: (?:(?:from|fra)\s*:\s*)(?:<\w+\s*class=""\w+\s*_?-?\s*sendername"">)?(?:<[a-z0-9|\/<>?,./:@~;'#!""£$%^&*()\-\-\+= ]+""\s*>)?([a-z,.:;\-_() ]+)(?:\[?mailto:)?(?:<)?'
    'EXPECTED RESULT: [1] Robert Radut, [2] Per Kaptain, [3] n/a, [4] Robert Radut'
    'SAMPLE TEXT: [1] From: Robert Radut [mailto:rr@webia.dk <mailto:rr@webia.dk> ]<br>, 
    '[2] From: <b class="gmail_sendername">Per Kaptain</b> <span dir="ltr">&lt;<a href="mailto:pka@nyfors.dk">pka@nyfors.dk</a>&gt;</span><br>Date: 2015-08-26 15:02 GMT+02:00<br>Subject: Bruger rettigheder ?<br>To: Robert Radut - Webia &lt;<a href="mailto:rr@webia.dk">rr@webia.dk</a>'
    '[4] Fra:</span></b><span style="font-size:10.0pt;font-family:&quot;Tahoma&quot;,&quot;sans-serif&quot;"> Robert Radut - Webia [mailto:<a href="mailto:rr@webia.dk"
    'TEST LINKS: 1: y, 2: y , 3: n/a , 4: y  '


    '#6 
    'VARIABLE NAME: subject 
    'REGEX CODE: (?>\s*<\s*\w+\s*>\s*(?:subject|Emne)\s*:\s*)(?>\s*re\s*:\s*|<\/?\w+>\s*)?([0-9a-z\/|?,.@'~#!""£$%^&*()\[\]{}\-_\+= ]+)(?:<\/?\w+>)?' 
    'EXPECTED RESULT: [1] zxzxzxzx, [2] Bruger rettigheder ?, [3] n/a, [4] TMS' 
    'SAMPLE TEXT: [1] <br>Subject: Re: zxzxzxzx, [2] <br>Subject: Bruger rettigheder ?<br>, [3] n/a, [4] <br><b>Emne:</b> TMS<u></u>'
    'TEST LINKS: 1: y, 2: y, 3: n/a, 4: y '


    '#7 
    'VARIABLE NAME: toEmail 
    'REGEX CODE [1]: (?:from\s*:\s*.*\n*\*?)?(?:<\s*\w+\s*\/?>)?(?:to\s*:\s*)(?:.*\n*.*?<\s*mailto\s*:\s*)?(?:[a-z0-9\/|?\-_!""£$%^&*()\+ ]+(?:<a\s*href=""mailto\s*:\s*|<\w+\s*\w+=""mailto\s*:\s*))?(\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b)(?>>\s*<\s*br\s*>|"">)'
    'REGEX CODE [2] - searches implicitly, where there is no "to:" provided: (?:(19|20)\d{2}\/?-?\d{2}\/?-?\d{2}\s*)?(?:\w+\+\d{2}\s*:\s*\d{2}\s*[a-z,.:;\-_ ]+\s*)(?:[0-9a-z\/|<>?:@~;'#!""£$%^&*()\-_\+= ]+)?(?:<\s*\w+\s*\w+="")?(?:\[?mailto\s*:\s*)(\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b)'
    'EXPECTED RESULT: [1] skk@demo.dk, [2] rr@webia.dk, [3] n/a, [4] jsc@bredbaandnord.dk'
    'SAMPLE TEXT: [1] To: Kevin D; Amanjot Kaur
    'Hej!<br><br>Vedhæftet opdateret tidsplan for projektet samt en liste over ting til afklaring. Lads os gennemgå det på mødet i morgen!<br><br>Var det noget med at Azri ville lave et sharepoint sted til alle filerne?<br><br>Inussiarnersumik inuulluaqqusillunga - Best Regards - Venlig hilsen<br><br>[cid:image001.jpg@01D05FF3.7ECFBEA0]<br><br>- E-Mail : skk@demo.dk<mailto:skk@demo.dk><br>Telefon: +45 99 00 00 00 - <br><br><hr>'
    '[2] To: Robert Radut - Webia &lt;<a href="mailto:rr@webia.dk">, [3] n/a, 
    '[4] 2015-08-26 12:17 GMT+02:00 Jacob Schou <span dir="ltr">&lt;<a href="mailto:jsc@bredbaandnord.dk"
    'TEST LINKS: 1: y, 2: y, 3: n/a, 4: y (using regex code [2])'


    '#8 
    'VARIABLE NAME: toPerson 
    'REGEX CODE: (?:til\s*:\s*(?:<\/?\w+>)?\s*|to\s*:\s*|hej\s*)([a-z\/,.:;'!""()*\-_\+ ]+)(?:\n+|<\/?\w+>|&lt;<\w+\s*\w+=""mailto|<o:p>|<\w:\w>|<\/?[a-z0-9<>?:@~;'#!""£$%^&*()\-_\+= ]+>)'
    'EXPECTED RESULT: [1] Kevin, [2] Robert Radut - Webia, [3] Robert, [4] Jacob Schou
    'SAMPLE TEXT: [1] <br>To: Kevin<br>, [2] <br>To: Robert Radut - Webia &lt;<a href="mailto:, [3] <p class="MsoNormal">Hej Robert<o:p></o:p>, [4] <b>Til:</b> Jacob Schou<br><br>
    'TEST LINKS: 1: y , 2: y, 3: y, 4: y '

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '***TEST URLS***:
        'http://nemnem.dk/webservice/mailtest.html'
        'http://nemnem.dk/webservice/mailtest2.html'
        'http://www.nemnem.dk/webservice/mailtest3.html'
        'http://www.nemnem.dk/webservice/mailtest4.html'

        '***TYPE IN TEST URL HERE***'
        Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("http://nemnem.dk/webservice/mailtest.html")
        Dim response As System.Net.HttpWebResponse = request.GetResponse
        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
        Dim sourcesample As String = sr.ReadToEnd
        Dim text As String = sourcesample


        '***ENTER ONE OF THE ABOVE REGULAR EXPRESSIONS INBETWEEN THE "" BELOW***'
        Dim pattern As String = "(?:til\s*:\s*(?:<\/?\w+>)?\s*|to\s*:\s*|hej\s*)([a-z\/,.:;'!""()*\-_\+ ]+)(?:\n+|<\/?\w+>|&lt;<\w+\s*\w+=""mailto|<o:p>|<\w:\w>|<\/?[a-z0-9<>?:@~;'#!""£$%^&*()\-_\+= ]+>)"
        Dim r As Regex = New Regex(pattern, RegexOptions.IgnoreCase)



        'Match the regular expression pattern against the contents of sourcesample - the HTML code in supplied URL.'
        'The code below tests all groups within a regular expression, returning matches from capturing groups, but not from non-capturing groups.'
        Dim m As Match = r.Match(text)
        Dim matchcount As Integer = 0
        Do While m.Success
            matchcount += 1
            ListBox1.DataSource = Nothing
            ListBox1.Items.Add("Match" & (matchcount))
            Dim i As Integer
            For i = 1 To 3
                Dim g As Group = m.Groups(i)
                ListBox1.Items.Add("Group" & i & "='" & g.ToString() & "'")
                Dim cc As CaptureCollection = g.Captures
                Dim j As Integer
                For j = 0 To cc.Count - 1
                    Dim c As Capture = cc(j)
                    ListBox1.Items.Add("Capture" & j & "='" & c.ToString() _
                     & "', Position=" & c.Index)
                Next
            Next
            m = m.NextMatch()
        Loop
    End Sub




    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class
