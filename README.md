<div align="center">

## \[A\+ Randomize Three Intro Art Splash Screen's Before Load\]


</div>

### Description

What this does is, Takes Three Variables/Possiblities and Randomizes them. It will randomizes three case's, case1, case2, case3, case one will show the intro splash screen art one, case two will show the intro splash screen art two, case three will show the intro splash screen art three. Inside each of the Intro's is coding that either on mouseclick, or in a timer will show form1/main, and hide the intro. This gives a nice effect to your intro art. The purpose is so that when you have intro art on your program, instead of having the same boring Splash Screen show up everytime...it randomizes between three forms that you have created...example in a .zip will be provided soon.

Words to help this search out: Intro, Intro Art, Splash Screen, Random, Randomize, Intro Splash Screen Art, Art, Screen, Windows, Case, Hack,
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[source](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/source.md)
**Level**          |Beginner
**User Rating**    |3.4 (24 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/source-a-randomize-three-intro-art-splash-screen-s-before-load__1-14270/archive/master.zip)





### Source Code

```
'This Code is to be inserted in the Main Ifrace of your program, it will be hidden while it randomizes which intro splash screen to load.
Private Sub Form_Load()
'main is the form that shows up after the intro art, could be form1, main or whatever you call it.
'main.visible=false , that hides the original form
Main.Visible = False
'makes the following coding randomized
Randomize
'lets it know there will be three options
Select Case Int((Rnd * 3) + 1)
'if case 1 is selected then the intro1 form loads
Case 1
intro1.Visible = True
'if case 2 is selected then the intro2 form loads
Case 2
intro2.Visible = True
'if case 3 is selected then the intro3 form loads
Case 3
intro3.Visible = True
'ends it
End Select
'*******END CODING*********
'inside intro1,2,3, you should have either a timer and after an alloted time it makes main.visible=true, and/or on MouseClick of the picture itself. so heres and example
'Intro1
'if someone clicks on the intro picture on intro1
'Private Sub Picture_click()
'hides the intro
'intro1.visible=false
'shows main iface
'main.visible=true
'End Sub
```

