<div align="center">

## A clean way to add control arrays at runtime\!


</div>

### Description

This code, will add elements to a control array element at runtime. This is very easy to do, as a control array is actually a collection and gives us a count.
 
### More Info
 
First off, make SURE you do this trick to VB. copy your control on the form, answer YES to create a control array, delete the new instance. Otherwise the control is not a collection.

New control is EXACTLY where the orginal was.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Trey Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/trey-smith.md)
**Level**          |Intermediate
**User Rating**    |2.8 (17 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/trey-smith-a-clean-way-to-add-control-arrays-at-runtime__1-6852/archive/master.zip)





### Source Code

```
'Load an new instance of myControl
'the index is the count, thus making it one
'greater than the prior index as index is 0 based
'and count starts at 1
Load pbVI(pbVI.Count)
'Count is now 1 greater so to address the control
'you just created reference count -1
With pbVI(pbVI.Count - 1)
'.Left = 100
'.Top = 600
'.Visible = True
End With
```

