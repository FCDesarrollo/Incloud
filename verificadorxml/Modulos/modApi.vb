Imports System.IO
Imports System.Net
Imports System.Text

Module modApi
    Public Const urlAPi As String = "http://apicrm.dublock.com/public/" '"http://localhost/apifc/public/" '

    Public Function ConsumeAPI(ByVal aMetodo As String, ByVal aDatos As String) As String
        Dim s As HttpWebRequest
        Dim enc As UTF8Encoding
        Dim response As HttpWebResponse
        Dim reader As StreamReader
        Dim rawresponse As String
        Dim postdata As String
        Dim postdatabytes As Byte()
        System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Try
            s = HttpWebRequest.Create(urlAPi & aMetodo)
            System.Net.ServicePointManager.UseNagleAlgorithm = False
            System.Net.ServicePointManager.Expect100Continue = False
            s.AllowReadStreamBuffering = False
            enc = New System.Text.UTF8Encoding()
            postdata = aDatos
            postdatabytes = enc.GetBytes(postdata)
            s.Method = "POST"
            's.ContentType = "application/x-www-form-urlencoded"
            s.KeepAlive = True
            s.ContentType = "application/json"
            's.Headers.Add("Content-Type", "appication/json")
            s.ContentLength = postdatabytes.Length
            s.Timeout = System.Threading.Timeout.Infinite
            's.ContinueTimeout = System.Threading.Timeout.Infinite
            's.ReadWriteTimeout = System.Threading.Timeout.Infinite
            Using stream = s.GetRequestStream()
                stream.Write(postdatabytes, 0, postdatabytes.Length)
            End Using

            response = s.GetResponse()
            reader = New StreamReader(response.GetResponseStream())

            rawresponse = reader.ReadToEnd()
            ConsumeAPI = rawresponse
        Catch ex As Exception
            'MsgBox("Error al Consumir API", vbExclamation, "Validación")
            ConsumeAPI = "false"
        End Try

    End Function
End Module
