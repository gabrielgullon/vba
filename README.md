# vba

### This is a collection of my most used subs and functions from when I used to code a lot in Excel and VBA

Since I need some place safe to save it and parts of this code I got from friends, colleagues and the internet, it's only fair that it gets public stored, so anyone can use it (if you can understand it lol).

### Modules (\*.bas)

- **A1_Aux.bas**: general-use scripts for daily _boring_ jobs
  <br>
- **A3_Collections.bas**: if you still use Excel to data analysis, these Subs can help you. Pretty much a dictionary, so that you don't need to use standard Excel functions... if you speak portuguese or read the few comments there, I'm sure you can get through it!
  <br>
- **A4_Combinacao.bas**: does a permutation/combination of every cell on each of the input cols and paste it on target cols

**Good to know**
The code bellow gets the last line on the adjacent cells to L3

> Dim last_line As Long
> last_line = ActiveSheet.Range("L3").CurrentRegion.Rows.Count

- It needs some attention, since it gets the content on the _region_, i.e., if you have something on Range("L3:L5") and Range("M4:M6") you won't get only the lenght of the content on col "L", but the sum of the lenght of L and M.
- Despite all that, it's still very usefull

### Future improvments

- good commenting and documentation
- improve formatting

  <br> <br> <br>

  > So understand
  > Don't waste your time always searching for those wasted years
  > Face up, make your stand
  > Realize you're living in the golden years

_Iron Maiden - Wasted Years_
