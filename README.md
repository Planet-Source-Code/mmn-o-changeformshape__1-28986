<div align="center">

## ChangeFormShape

<img src="PIC200111261418416243.gif">
</div>

### Description

The code changes the form shape to either 'Elliptic', 'Rounded Rectangle', 'Rectangle', 'Polygon' or 'Picture-Shaped'.

In the code you can also see how to COMBINE different shapes, so you can create a very neat interface.

These functions are in my code:

1. Change Form Shape (See above)

2. Nice interface (somewhat like XP I believe)

3. Move and Resize a border-less form (Borderstyle = '0 - None')

4. On Resize - align controls (class module)

5. Text in Pictureboxes

6. EasyMove forms

7. Change Picture Shape

There aren't many comments in the source, but it doesn't have to be either, even a "newbie" will understand what's going on...

Before the code only showed how to change the form's shape, now it also shows how to change shapes of picture-boxes.

In fact, I believe you can change the shape of everything that has a hWnd variable...

I haven't tried it, but I think you can show oval videos too!!! (or polygon or whatever you wish...)

Please vote...

<

----

>

Anyway, here's an FAQ:

Q: Is it simple?

A: Yes!

----

Q: How simple?

A: Very...

----

Q: Is this Windows XP or VB.NET?

A: Nope, I used Win98 and VB 6.0

----

Q: Is all this code made by you?

A: No, but the thanks are in the source.

----

Q: Are you the best programmer in the world?

A: Who, me?

<

----

>

Thank you for reading this - now download, try out and vote!

(If there's anything that's not working or you wish me to add, please let me know!)

/Mikael Nordfelth

Picture-shaped-form has been renewed, now it's not Niknak's code... (it was quite slow compared to this)
 
### More Info
 
The form will not be restored until the next restart...

To (almost) fix this, to like this:

In Form_Resize you put a CreateRectangleRgn that is the size of Me.Width and Me.Height


<span>             |<span>
---                |---
**Submitted On**   |2001-11-26 18:29:46
**By**             |[MMN\-o](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mmn-o.md)
**Level**          |Intermediate
**User Rating**    |4.6 (88 globes from 19 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[ChangeForm3758611262001\.zip](https://github.com/Planet-Source-Code/mmn-o-changeformshape__1-28986/archive/master.zip)








