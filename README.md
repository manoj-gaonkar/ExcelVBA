```
'  --> this is commment in VBlearn
```

# Intro
```
Sub Macro_name()

' code here

End Sub
```
# Sheets
```
Sheets("Sheet1").Activate
Sheets("Sheet1").Select
```


# Workbooks
```
Workbooks("VBlearn1.xlsx").Activate
```

# Display MsgBox
```
MsgBox (ThisWorkbook.Name)
MsgBox (ActiveWorkbook.Name)
```

# Variables
```
name = "YourName"

'write names to some cells
Range("c1:c4").value = name
```

# For loop

### Example 1
```
Dim x As Integer
For x = 1 to 5
    MsgBox(x)
Next x
```

### Example 2
```
Dim y As Integer

For y = 1 To 20 Step 3 #here step 3 is like step skip
    Cells(y, 1).Value = y
Next y
```


### Example 3 / for loop in reverse order
```
Dim y as Integer
Dim y As Integer
For y = 10 To 1 Step -1
    Cells(y, 1).Value = 21 - y
Next y
```

### Example 4 / for loop to display sheetsname
```
For y = 1 To ThisWorkbook.Sheets.Count
    MsgBox (Sheets(y).Name)
Next y
```

# For Each / Next
```
Dim sht as Worksheet

For each sht in ThisWorkbook.Sheets
    Msgbox sht.name
Next sht
```