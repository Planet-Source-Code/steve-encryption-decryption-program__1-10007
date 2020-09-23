<div align="center">

## Encryption/decryption program


</div>

### Description

This article explains how to encrypt a message and then decrypt it using an array and loops. Using random letters to change the original letters of the original message.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-encryption-decryption-program__1-10007/archive/master.zip)





### Source Code

```
First let me explain the algorithm.
For encrypting:
1) Dimension variables
2) Clear all variables
3) get the message
4) Loop through the message
5) In the loop, firstly randomize a number from 1 to 110 and hold this number in an array. Secondly increment a value by 1. Now, get the new character and add it to the encrypted message(Using chr$ in VB).
Now after the loop has finished the message has been encrypted.
To Decrypt it:
1) Get the coded message.
2) Loop through teh coded message.
3) increment a value by 1.
4) by using the Chr$, mid$ and asc functions, take away the ascii relative to codded letter and add the letter to a variable (decoded message).
In a module set these variables:
Global a% 'the value that increments in the loops
Global msgnum(10000) As Long 'the array that holds the ascii numericals relative to the letters
Global codedmsg As String 'the encrypted message
In your encryption procedure:
Dim n%
Dim x%
Dim message$
Dim emessage$
Dim word$
Dim ctext
Dim i As Integer
Erase msgnum 'erase all value in the array
a% = 0
ctext = txtmessage.Text
word$ = ""
For i = 1 To Len(ctext) 'loop through the string
 Randomize
  x% = Int((110 * Rnd) + 1) 'randomize a number from 1 to 110.
  a% = a% + 1 'Increment this value as it is used in the array
 word$ = word$ & Chr$(Asc(Mid(c_text, i, 1)) + x%)
'add the original letter ascii value to the randomized value and produce that character
  msgnum(a%) = x% 'hold the randomized number in the array
Next i
codedmsg = word$
In your decryption procedure:
Dim msg$
Dim x%
Dim ctext
Dim word$
Dim decodedmsg
Dim i As Integer
ctext = codedmsg
word$ = ""
a% = 0
For i = 1 To Len(c_text) 'Loop through the coded message
  a% = a% + 1 'Increment value by 1
word$ = word$ & Chr$(Asc(Mid(c_text, i, 1)) - msgnum(a%))
'this time take away the the randomized value, which is held in the array from the codded character ascii value to produce the original letter
Next i
decodedmsg = word$ 'Decoded msg
```

