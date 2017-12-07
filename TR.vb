Sub GrabLastNames()

    'dimension (set aside memory for) our variables
    Dim objIE As InternetExplorer
    Dim ele As Object
    Dim y As Integer
    
    'start a new browser instance
    Set objIE = New InternetExplorer
    'make browser visible
    objIE.Visible = True
    
    'navigate to page with needed data
    objIE.navigate "http://names.mongabay.com/most_common_surnames.htm"
    'wait for page to load
    Do While objIE.Busy = True Or objIE.readyState <> 4: DoEvents: Loop
    
    'we will output data to excel, starting on row 1
    y = 1
    
    'look at all the 'tr' elements in the 'table' with id 'myTable',
    'and evaluate each, one at a time, using 'ele' variable
    For Each ele In objIE.document.getElementById("myTable"). _
      getElementsByTagName("tr")
        'show the text content of 'tr' element being looked at
        Debug.Print ele.textContent
        'each 'tr' (table row) element contains 4 children ('td') elements
        'put text of 1st 'td' in col A
        Sheets("Sheet1").Range("A" & y).Value = ele.Children(0).textContent
        'put text of 2nd 'td' in col B
        Sheets("Sheet1").Range("B" & y).Value = ele.Children(1).textContent
        'put text of 3rd 'td' in col C
        Sheets("Sheet1").Range("C" & y).Value = ele.Children(2).textContent
        'put text of 4th 'td' in col D
        Sheets("Sheet1").Range("D" & y).Value = ele.Children(3).textContent
        'increment row counter by 1
        y = y + 1
    'repeat until last ele has been evaluated
    Next

    'save the Excel workbook
    ActiveWorkbook.Save

End Sub
