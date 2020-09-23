<div align="center">

## \_\_ Dynamically allocate arrays when you don't know how many items to store\!


</div>

### Description

All programmers need to allocate arrays to store data, and very often they don't know how much they will be storing. here is a beginners tutorial that shows how to dynamically allocate an array on the fly that will only allocate as many items as are needed!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Fosters](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fosters.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fosters-dynamically-allocate-arrays-when-you-don-t-know-how-many-items-to-store__1-43726/archive/master.zip)





### Source Code

<P><FONT face="Courier New" size=2>This is a short tutorial on dynamically building arrays <br>
(with examples for 1 and 2 dimensions).<br>
There are many occasions where you need <br>
to allocate an array, but don't know what the upper bounds are.<br>
Shown here is an efficient tried and trusted method.<br>
The whole concept revolves around UBOUND - the upper limit of your array.<br>
Knowing the upper limit allows you to increase it's size by as much as <br>
you need to, without having to initially allocate a huge <br>
array at the start! <br><br>
The key points are <br>
'define a 0 bounded array, so that redims later on do not fail<br>
ReDim sTempArray(0) <br>
'perform your loop to work out what must go in each element of your array <br>
Do <br>&nbsp;&nbsp;&nbsp; If we need to allocate another item to the array Then&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'redimension the array to accomodate the new data&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ReDim Preserve sTempArray(UBound(sTempArray) + 1)&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
sTempArray(UBound(sTempArray) - 1) = ???&nbsp; <br>&nbsp;&nbsp;&nbsp;
End If <br>
Loop Until ??? <br><br>
'this method allocates 1 too many array items, so reduce it by 1 <br>
ReDim Preserve sTempArray(UBound(sTempArray) - 1) <br>
'Array is ready to be returned with only the data you have allocated<br><br>
You can paste the following example code to an app. <br>
run your app, Pause it, and in the immediates window, type GetArrayData <br><br>
Sub GetArrayData()&nbsp; <br>
Dim sRecieve1DArray() As String <br>
Dim sRecieve2DArray() As String&nbsp; <br><br>&nbsp;&nbsp;&nbsp;
sRecieve1DArray = ReturnOneDimensionalArray&nbsp; <br>&nbsp;&nbsp;&nbsp; Debug.Print UBound(sRecieve1DArray) Debug.Print sRecieve1DArray(0), sRecieve1DArray(1), sRecieve1DArray(2)&nbsp; <br><br>&nbsp;&nbsp;&nbsp; sRecieve2DArray = ReturnTwoDimensionalArray&nbsp; <br>&nbsp;&nbsp;&nbsp;
Debug.Print UBound(sRecieve2DArray, 2)&nbsp; <br>&nbsp;&nbsp;&nbsp; Debug.Print sRecieve2DArray(0, 0), sRecieve2DArray(0, 1), sRecieve2DArray(0, 2)&nbsp; <br>&nbsp;&nbsp;&nbsp;
Debug.Print sRecieve2DArray(1, 0), sRecieve2DArray(1, 1), sRecieve2DArray(1, 2) <br>
End Sub <br><br>
Function ReturnOneDimensionalArray() As String() <br>
Dim sTempArray() As String <br>
Dim iCount As Integer&nbsp; <br><br>&nbsp;&nbsp;&nbsp; 'initially define the array otherwise the other redims will fail&nbsp; <br>&nbsp;&nbsp;&nbsp; ReDim sTempArray(0)&nbsp; <br>&nbsp;&nbsp;&nbsp; iCount = 0&nbsp; <br><br>&nbsp;&nbsp;&nbsp; Do&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'redimension the array to the upper limt + 1&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
ReDim Preserve sTempArray(UBound(sTempArray) + 1)<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'populate into the upper limit -1&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
sTempArray(UBound(sTempArray) - 1) = Chr(65 + iCount)&nbsp; <br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; iCount = iCount + 1&nbsp; <br>&nbsp;&nbsp;&nbsp; Loop Until iCount &gt;=
  26<br><br>&nbsp;&nbsp;&nbsp; 'you have 1 more index than necessary, so reduce it by 1&nbsp; <br>&nbsp;&nbsp;&nbsp; ReDim Preserve sTempArray(UBound(sTempArray) - 1)&nbsp; <br><br>&nbsp;&nbsp;&nbsp; 'assign the temporary array to the function for return&nbsp; <br>&nbsp;&nbsp;&nbsp;
ReturnOneDimensionalArray = sTempArray <br>
End Function <br><br>
Function ReturnTwoDimensionalArray() As String() <br>
Dim sTempArray() As String <br>
Dim iCount As Integer&nbsp; <br><br>&nbsp;&nbsp;&nbsp; 'initially define the array otherwise the other redims will fail&nbsp; <br>&nbsp;&nbsp;&nbsp; 'remember, you can only redim the last dimension&nbsp; <br>&nbsp;&nbsp;&nbsp; ReDim sTempArray(2, 0)&nbsp; <br><br>&nbsp;&nbsp;&nbsp;
iCount = 0&nbsp; <br>&nbsp;&nbsp;&nbsp; Do&nbsp; <br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'redimension the array to the upper limt + 1&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'you are referencing and increasing the 2nd dimension&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ReDim Preserve sTempArray(2, UBound(sTempArray, 2) + 1)&nbsp; <br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'populate into the upper limit -1&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; sTempArray(0, UBound(sTempArray, 2) - 1) = Chr(65 + iCount)&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; sTempArray(1, UBound(sTempArray, 2) - 1) = Chr(97 + iCount)&nbsp; <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; iCount = iCount + 1&nbsp; <br>&nbsp;&nbsp;&nbsp; Loop Until iCount &gt;=
  26&nbsp; <br><br>&nbsp;&nbsp;&nbsp;
'you have 1 more index than necessary (on the 2nd dimension), so reduce it by 1&nbsp; <br>&nbsp;&nbsp;&nbsp; ReDim Preserve sTempArray(2, UBound(sTempArray, 2) - 1)&nbsp; <br>&nbsp;&nbsp;&nbsp;
'assign the temporary array to the function for return&nbsp; <br>&nbsp;&nbsp;&nbsp;
ReturnTwoDimensionalArray = sTempArray <br>End Function</FONT>
 </P>

