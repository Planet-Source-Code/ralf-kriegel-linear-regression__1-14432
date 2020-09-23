<div align="center">

## linear regression


</div>

### Description

Calculation of a linear regression

(straight line) using a X,Y data array
 
### More Info
 
X,Y data array

requires a minimum of two data points

Slope and Y-intersection


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ralf Kriegel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ralf-kriegel.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ralf-kriegel-linear-regression__1-14432/archive/master.zip)





### Source Code

```
Public Sub LinRegsngArr(ByRef sng_XArray!(), ByRef sng_YArray!(), _
  ByVal lowLimInd&, ByVal uppLimInd&, ByRef Slope!, ByRef YSection!)
'this Code calculates a linear regression
'using the points of two Arrays (X,Y)
'and gives back slope and Y-intersection
'of the straight line
Dim sng_XSum!, sng_YSum!, sng_XQuad!, sng_YQuad!, sng_XYProd!, sng_Fract!
Dim lng_Index&, ValuesCounts&, sng_Zaehler!
 ValuesCounts = uppLimInd - lowLimInd + 1
  For lng_Index = lowLimInd To uppLimInd
   sng_XSum = sng_XSum + sng_XArray(lng_Index)
   sng_YSum = sng_YSum + sng_YArray(lng_Index)
   sng_XQuad = sng_XQuad + sng_XArray(lng_Index) ^ 2
   sng_YQuad = sng_YQuad + sng_YArray(lng_Index) ^ 2
   sng_XYProd = sng_XYProd + sng_YArray(lng_Index) * sng_XArray(lng_Index)
  Next
 sng_Fract = ValuesCounts * sng_XQuad - sng_XSum ^ 2
 Slope = (ValuesCounts * sng_XYProd - sng_XSum * sng_YSum) / sng_Fract
 YSection = (sng_YSum * sng_XQuad - sng_XSum * sng_XYProd) / sng_Fract
End Sub
```

