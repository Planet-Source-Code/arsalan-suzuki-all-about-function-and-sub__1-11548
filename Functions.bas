Attribute VB_Name = "Functions"
'If you include this then any undeclared
'variable generates error
Option Explicit

'###################################
'# This is what you need to become #
'# a master, in using function or  #
'# sub.                            #
'# If you have any comment pl mail #
'# me                              #
'# if you think there is something #
'# wrong in this tutorial then pl  #
'# inform me ASAP                  #
'###################################

'Point to be noted::::------->
'Same method also applies for Sub


' This is new trick (I havnt used in my prev tutorial)
' I think you are familiar with this.
' If you set the Start up object to Sub Main
' (which can be found on project properties)
' Then before starting your application
' main() sub will be called.
Sub Main()
Form1.Show ' show Form1
End Sub

'From MSDN documentation (only this bit)
'This is the syntax of functino or sub
'##########################
'[Public | Private | Friend] [Static] Function (or Sub) name [(arglist)] [As type]
'[statements]
'[name = expression]
'[Exit Function]
'[statements]
'[name = expression]
'End Function (or Sub)
'##########################
'By default scope of the function is public

'About the scope of the function
'--------------------------------
'Public Indicates that the Function procedure is accessible to all other procedures in allmodules. If used in a module that contains an Option Private, the procedure is not available outside theproject.
'Private Indicates that the Function procedure is accessible only to other procedures in the module where it is declared.
'Friend  Used only in a class module. Indicates that the Function procedure is visible throughout the project, but not visible to a controller of an instance of an object.
'Static Indicates that the Function procedure's localvariables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Function, even if they are used in the procedure.
'--------------------------------

'The arglist argument has the following syntax :
'[Optional] [ByVal | ByRef] [ParamArray] varname[( )] [As type] [= defaultvalue]
'By default the argument is passed with ByVal

'I think you have understood upto this much
'--------------------------------
'I'll take you step by step.

'---------------ByVal---------------------
'In this function the value of the variable is passed
'(only value)
'suppose we called ChangeCaption(k)
'so k="ABCDEFGHIJ"
'and only ABCDEFGHIJ is passed
Public Function ChangeCaption(ByVal FormsCaption As String)
ChangeCaption = FormsCaption
End Function

'---------------ByRef---------------------
'This function is public (as its not defined so default is used).

'You will be wondering without returning the value
'the value of the variable has changed.

'This is because we passed the address of TextToChange.

'Little about the Variable
'-------------------------
'Each variable in VB or any lnaguage, is stored
'in different addresses. Address are usually random,
'so you cannot predict in which address its been stored.
'suppose in this e.g we pass the address
'of the variable TextToChange (suppose its stored in
'4002),so we are passing 4002 address to the function.
'As 4002 address contains the value of the variable
' 't' (see Form1), any change in TextToChange will
'also change the variable 't'
'So............
Function TextChange(ByRef TextToChange As String)
TextToChange = "Hi, I'm also changed"
End Function

'I think you have understood upto this bit

'---------------Optional----------------------
'Its very simple
'Text variable is declared as optional one
'If its not given in the function (see Form1)
'it will display the default value
'Other wise it will act as normal variable
Public Function Hi(Optional Text As String = "Optional you don't have to enter")
MsgBox Text
End Function

'---------------ParamArray--------------------
'In paramarray you can pass multiple parameter
'and each parameter will be stored in the
'array Numbers
'so
'1st parameter = Numbers(0)
'2nd parameter = Numbers(1)
'3rd parameter = Numbers(2)
'and so on
Public Function Sum(ParamArray Numbers())
Dim temp As Integer
Dim i As Integer
temp = 0

For i = LBound(Numbers) To UBound(Numbers)
temp = temp + Numbers(i)
Next
Sum = temp
End Function

'---------------------------------------------

'Well this is end of the tutorial
'I hope you have learned something new
'By using this concept you will be
'able to make more sophisticated functions
