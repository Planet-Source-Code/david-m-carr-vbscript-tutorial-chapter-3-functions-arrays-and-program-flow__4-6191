<div align="center">

## VBScript Tutorial: Chapter 3\-\-Functions, Arrays, and Program Flow


</div>

### Description

The complete lowdown on creating functions, declaring and using arrays to store data, as well as an intro to program flow statements().
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David M\. Carr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-m-carr.md)
**Level**          |Beginner
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-m-carr-vbscript-tutorial-chapter-3-functions-arrays-and-program-flow__4-6191/archive/master.zip)





### Source Code

```
<hr>
<h6><font face="Verdana">Note: In syntax definitions, italics means that is
something to be filled in, &quot;[]&quot; mean that it is optional, and
&quot;...&quot; mean that there can be more than one.</font></h6>
<hr>
<p><font face="Verdana">To declare variables, you use the Dim keyword.&nbsp; The
syntax for Dim is the following:</font><code></p>
<div class="vbscode">
 <p><font face="Verdana">Dim </font></code><font face="Verdana"><em>identifier</em></font></p>
</div>
<p><font face="Verdana">You can declare multiple variables with Dim by
separating them with commas.</font><code></p>
<div class="vbscode">
 <p><font face="Verdana">Dim <em>indentifier1</em>, <em>identifier2</em>, <em>identifier3</em>
 ...</font></code></p>
</div>
<p><font face="Verdana">An identifier starts with a letter, and then can contain
up to 254 characters. &nbsp; Generally it is a good idea to stick to letters,
numbers, and underscores, but some other punctuation marks are allowed.&nbsp;
Certain words cannot be used for identifiers, as they are used by VBScript.&nbsp;
These are called <em>reserved words</em>.&nbsp; Dim is optional in most cases
(but not all), since the browser will define the variable when you first use it.</font></p>
<p><font face="Verdana">To set a value to a variable, you use the = operator.&nbsp;
It is used for two things, comparison and assigning values.&nbsp; For example:</font></p>
<div class="vbsCode">
 <code>
 <p><font face="Verdana">Dim x<br>
 x = 8<br>
 If x = 8 Then Alert &quot;X is 8&quot;<br>
 x = x + 9</font></code></p>
</div>
<p><font face="Verdana">This example defines X as a variable, sets the value 8,
shows a message if X is equal to 8, and then adds 9 to the current value of X.&nbsp;
Note that a variable can be used on both sides of the = operator.</font></p>
<hr>
<p><font face="Verdana">A subroutine is easy to define.&nbsp; All you do is
follow this syntax:</font></p>
<code>
<div class="vbscode">
 <p><font face="Verdana">Sub <em>identifier</em>([<em>parameter list</em>])<em><br>
 </em>&nbsp;&nbsp;&nbsp; *** VBScript Code Here ***<br>
 End Sub</font></code></p>
</div>
<p><font face="Verdana">All lines of code between the Sub and the End Sub will
be executed whenever the subroutine is called.&nbsp; The parameter list is
optional.&nbsp; It is a list of identifiers, separated by commas.&nbsp; After
the function syntax will be an example.</font></p>
<code>
<div class="vbscode">
 <p><font face="Verdana">Function <em>identifier</em>([<em>parameter list</em>])<br>
 &nbsp;&nbsp;&nbsp; *** VBScript Code Here ***<br>
 End Function</font></code></p>
</div>
<p><font face="Verdana">Notice the similarities.&nbsp; The only difference is
that Sub is replaced by Function. &nbsp; However, it is expected that in the
function body (the &quot;code here&quot; section), you will return a value.&nbsp;
You do this by using the function name in an assignment statement.&nbsp; For
example:</font></p>
<code>
<div class="vbscode">
 <p><font face="Verdana">Function Squared(aNumber)<br>
 &nbsp;&nbsp;&nbsp; Squared = aNumber * aNumber<br>
 End Function</font></code></p>
</div>
<p><font face="Verdana">(The asterisk is the VBScript operator for
multiplication.)</font></p>
<p><font face="Verdana">All variables in VBScript are variants.&nbsp; They can
hold one type of data in one place in the script, and another later on.&nbsp;
This code is valid:</font></p>
<code>
<div class="vbscode">
 <p><font face="Verdana">Dim X<br>
 X = 5<br>
 X = &quot;Hello&quot;<br>
 X = 8.35<br>
 X = True<br>
 X = Date()</font></code></p>
</div>
<p><font face="Verdana">Many programming languages have variables which must be
only one type, like integer, string, decimal, etc., but VBScript isn't like
that.&nbsp; Thus, in some cases, it is necessary to use built-in conversion
functions so that data is interpreted as the correct type.&nbsp; Some conversion
functions are CStr (to string), CSng (to floating point, single precision) and
CInt (to integer).</font></p>
<hr>
<p><font face="Verdana">Arrays are also defined with Dim.&nbsp; Here is the
syntax to declare an array.</font></p>
<div class="vbscode">
 <code>
 <p><font face="Verdana">Dim <em>identifier</em>([<em>subscript1</em>, <em>subscript2</em>
 ...])</font></code></p>
</div>
<p><font face="Verdana">From 0 to 60 subscripts are supported, though if you use
0 subscripts, you will have to use ReDim, which is discussed next,&nbsp; before
you can use the array.&nbsp; ReDim changes the size of the array, so that is
stores more or less data.&nbsp; The subscript is one less than the number of
items in the array, since the array always starts at 0.</font></p>
<div class="vbscode">
 <code>
 <p><font face="Verdana">ReDim [Preserve] <em>identifier</em>(<em>subscript1</em>,
 <em>subscript2</em> ...)</font></code></p>
</div>
<p><font face="Verdana">If used, Preserve saves the data currently in the array.&nbsp;
Otherwise, the array is emptied.&nbsp; If you ReDim the array to make it
smaller, data in the portion removed will always be lost.</font></p>
<p><font face="Verdana">Here is an example:</font></p>
<div class="vbscode">
 <code>
 <p><font face="Verdana">Dim AnArray(5) <span class="vbsComment">'Can hold 6
 values, AnArray(0) to AnArray(5)</span><br>
 AnArray(3) = &quot;Hello&quot; <span class="vbsComment">'Sets a value to the
 fourth space in the array</span><br>
 ReDim Preserve AnArray(4) <span class="vbsComment">'Value still there</span><br>
 ReDim AnArray(9, 2) <span class="vbsComment">'Value lost, didn't use Preserve</span><br>
 AnArray(3,5) = &quot;Goodbye&quot; <span class="vbsComment">'Set another value</span><br>
 ReDim Preserve AnArray(3, 1) <span class="vbsComment">'Data lost, end portion
 removed</span></font></code></p>
</div>
<p><font face="Verdana">Whenever you use the Dim statement to declare a variable
or array, the values are filled with either a 0 (if it's a number) or an empty
string.&nbsp; This is called initialization.</font></p>
<p><font face="Verdana">Where you place the Dim statement is very important
regardless of whether you are using it to declare a variable or array.&nbsp; The
placement of the Dim statement determines where the variable or array can be
used (the scope).&nbsp; If you put the Dim statement in a Script tag, but not in
a subroutine or function, it is declared 'globally', meaning that you can use
that variable in any function or subroutine on the page.&nbsp; If you put the
Dim statement in a function, it can only be used in that function, as with
subroutines.</font></p>
<code>
<div class="vbscode">
 <p><font face="Verdana">&lt;SCRIPT TYPE=&quot;text/vbscript&quot;
 LANGUAGE=&quot;VBScript&quot;&gt;<br>
 &lt;!--<br>
 Dim X <span class="vbsComment">'Declares X Globally</span><br>
 Sub Test1()<br>
 &nbsp;&nbsp;&nbsp; Dim Y <span class="vbsComment">'Declares Y in a subroutine</span><br>
 &nbsp;&nbsp;&nbsp; Y = 30 <span class="vbsComment">'Sets a value in Y</span><br>
 End Sub <span class="vbsComment">'Value in Y is lost at the end of the
 subroutine</span><br>
 Sub Test2()<br>
 &nbsp;&nbsp;&nbsp; Alert X + Y <span class="vbsComment">'Y is 0, since it
 wasn't declared globally or in this subroutine</span><br>
 End Sub<br>
 <br>
 Test1 <span class="vbsComment">'Calls the first subroutine</span><br>
 Test2 <span class="vbsComment">'Calls the second subroutine</span><br>
 '--&gt;<br>
 &lt;/SCRIPT&gt;</font></code></p>
</div>
<p><font face="Verdana">To make your coding easier to understand and to make it
easier to notice typos, you can put <span class="vbsCode">Option Explicit</span>
at the beginning of each Script tag, right after the start of the comment tag.&nbsp;
This causes a message to be displayed by the browser whenever a variable isn't
declared using Dim.</font></p>
<hr>
<p><font face="Verdana">In <a href="/vb/scripts/ShowCode.asp?lngWId=4&amp;txtCodeId=6189">Chapter
1</a>, two methods of putting script code into an HTML document were discussed,
both for event handlers.&nbsp; There is another way of writing event handlers,
which is writing a subroutine.&nbsp; The subroutine's name is composed of the
object's ID/name, followed by an underscore and the event.&nbsp; For the button
Button1, you would handle a click event like this:</font></p>
<code>
<div class="vbscode">
 <p><font face="Verdana">&lt;SCRIPT TYPE=&quot;text/vbscript&quot;
 LANGUAGE=&quot;VBScript&quot;&gt;<br>
 &lt;!--<br>
 Sub Button1_OnClick()<br>
 &nbsp;&nbsp;&nbsp; Alert &quot;This is a message.&quot;, vbExclamation, _<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &quot;VBScript Test&quot;<br>
 End Sub<br>
 '--&gt;<br>
 &lt;/SCRIPT&gt;</font></code></p>
</div>
<p><font face="Verdana">This is much easier than enclosing all the script code
inside the object tag, and is more compatible with current browsers than using
the For and Event attributes of the Script tag.&nbsp; This also allows all event
handlers to be put together in one script tag, so that they are easier to find
and reuse.</font></p>
<p><font face="Verdana">By now, you may have guessed what the final method of
writing VBScript in HTML is. &nbsp; It is writing code directly within a Script
tag.&nbsp; Any statements in a Script tag but not in a subroutine or function
are run when the browser reaches that point in the page.&nbsp; This allows you
to add content to the middle of a page easily.</font></p>
<p><font face="Verdana">The ampersand (&quot;&amp;&quot;) is used in VBScript to
attach two strings.</font></p>
<code>
<div class="vbscode">
 <p><font face="Verdana">&lt;SCRIPT TYPE=&quot;text/vbscript&quot;
 LANGUAGE=&quot;VBScript&quot;&gt;<br>
 &lt;!--<br>
 Document.Write &quot;This page was updated &quot; &amp; _<br>
 &nbsp;&nbsp;&nbsp; Document.LastModified &amp; &quot;.&lt;BR&gt;&quot;<br>
 Document.Write &quot;It was &quot; &amp; Time() &amp; _<br>
 &nbsp;&nbsp;&nbsp; &quot; when you loaded this page.&quot;<br>
 '--&gt;<br>
 &lt;/SCRIPT&gt;</font></code></p>
</div>
<p><font face="Verdana">Below is what the above script would output.</font></p>
<p><font face="Verdana"><script language="VBScript" type="text/vbscript"><!--
Document.Write "This page was update " & Document.LastModified & ".<BR>"
Document.Write "It was " & Time() & " when you loaded this page."
--></script>
 This page was update 05/25/2000 16:18:48.<br>
It was 4:18:48 PM when you loaded this page.</font></p>
<hr>
<p><font face="Verdana">Boolean values are either true or false.&nbsp;
Internally, VBScript always uses -1 for true and 0 for false, though when it is
converting numbers to boolean values, all non-zero numbers are true.&nbsp; Other
than converting number to boolean values, you can also get boolean values by
using comparison operators or logical operators.</font></p>
<table border="1">
 <tbody>
  <tr>
   <td><font face="Verdana">Comparison Operator</font></td>
   <td><font face="Verdana">Function</font></td>
   <td><font face="Verdana">Logical Operator</font></td>
   <td><font face="Verdana">Function</font></td>
  </tr>
  <tr>
   <td><font face="Verdana"><em>val1</em> = <em>val2</em></font></td>
   <td><font face="Verdana">True if they are the same</font></td>
   <td><font face="Verdana"><em>val1</em> And <em>val2</em></font></td>
   <td><font face="Verdana">True if both are true</font></td>
  </tr>
  <tr>
   <td><font face="Verdana"><em>val1</em> &lt;&gt; <em>val2</em></font></td>
   <td><font face="Verdana">True if they are not the same</font></td>
   <td><font face="Verdana"><em>val1</em> Or <em>val2</em></font></td>
   <td><font face="Verdana">True if one or both are true</font></td>
  </tr>
  <tr>
   <td><font face="Verdana"><em>val1</em> &lt; <em>val2</em></font></td>
   <td><font face="Verdana">True if val1 is less than val2</font></td>
   <td><font face="Verdana">Not <em>val</em></font></td>
   <td><font face="Verdana">True if val is false</font></td>
  </tr>
  <tr>
   <td><font face="Verdana"><em>val1</em> &gt; <em>val2</em></font></td>
   <td><font face="Verdana">True if val1 is greater than val2</font></td>
   <td><font face="Verdana"><em>val1</em> Xor <em>val2</em></font></td>
   <td><font face="Verdana">True if only one is true</font></td>
  </tr>
  <tr>
   <td><font face="Verdana"><em>val1</em> &lt;= <em>val2</em></font></td>
   <td><font face="Verdana">True if val1 is either less than or the same as
    val2</font></td>
   <td><font face="Verdana"><em>val1</em> Eqv <em>val2</em></font></td>
   <td><font face="Verdana">True if both are true or both are false</font></td>
  </tr>
  <tr>
   <td><font face="Verdana"><em>val1</em> &gt;= <em>val2</em></font></td>
   <td><font face="Verdana">True if val1 is either greater than or the same
    as val2</font></td>
   <td><font face="Verdana"><em>val1</em> Imp <em>val2</em></font></td>
   <td><font face="Verdana">True if val2 is true or both are false</font></td>
  </tr>
 </tbody>
</table>
<p><font face="Verdana">Conditional statements allow you to make a section of
code dependant on something else. There are four types of conditional statements
in VBScript, three of which use boolean conditions.</font></p>
<p><font face="Verdana">In the syntax definitions, statements can be any type of
VBScript statement.</font></p>
<div class="VBSCode">
 <code>
 <ol>
  <li><font face="Verdana">If <em>condition</em> Then <em>statement</em> [Else
   <em>statement</em>]</font>
  <li><font face="Verdana">If <em>condition</em> Then<br>
   <em>&nbsp;&nbsp;&nbsp; statement(s)<br>
   </em>[Else<br>
   &nbsp;&nbsp;&nbsp; <em>statement(s)</em>]<em><br>
   </em>End If</font>
  <li><font face="Verdana">If <em>condition</em> Then<br>
   &nbsp;&nbsp;&nbsp; <em>statement(s)</em><br>
   ElseIf <em>condition</em><br>
   &nbsp;&nbsp;&nbsp; <em>statement(s)</em><br>
   ElseIf <em>condition</em><br>
   &nbsp;&nbsp;&nbsp; <em>statement(s)<br>
   </em>ElseIf <em>condition</em><br>
   &nbsp;&nbsp;&nbsp; <em>statement(s)<br>
   </em>...<br>
   End If</font>
  <li><font face="Verdana">Select Case <em>expression</em><br>
   &nbsp;&nbsp;&nbsp; Case <em>value</em><br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <em>statement(s)</em><br>
   &nbsp;&nbsp;&nbsp; Case <em>value</em><br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <em>statement(s)</em><br>
   &nbsp;&nbsp;&nbsp; Case <em>value</em><br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <em>statement(s)</em><br>
   &nbsp;&nbsp;&nbsp; ...<br>
   &nbsp;&nbsp;&nbsp; [Case Else<br>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <em>statement(s)</em>]<br>
   End Select</font></li>
 </ol>
 </code>
</div>
<p><font face="Verdana">The first conditional statement is best suited for
situations where you only want to perform one action if the condition is true,
and possibly one if it is false.&nbsp; Ex:</font></p>
<div class="VBSCode">
 <p><code><font face="Verdana">If UserName = &quot;George&quot; Then Alert
 &quot;Hi, George.&quot; Else Alert &quot;You're not George.&quot;</font></code></p>
</div>
<p><font face="Verdana">The second conditional statement is best suited for
situations where you&nbsp; want to perform one or more actions if the condition
is true, and possibly one or more actions if it is not.</font></p>
<p><font face="Verdana">The third conditional statement is best suited for
situations where there are many conditions, and depending on them, you want to
take certain actions.</font></p>
<p><font face="Verdana">The fourth conditional statement is the oddball.&nbsp;
For Expression, you give anything that evaluates to a single value.&nbsp; For
each Value, , you provide a value, such as a string, a number, or boolean value.&nbsp;
It chooses one of the groups of statements after a Case Value pair by comparing
the expression to the values.&nbsp; Ex:</font></p>
<div class="VBSCode">
 <p><code><font face="Verdana">Dim aVar<br>
 Randomize <span class="vbscomment">'Insure that the numbers are truly random</span><br>
 aVar = Int(3 * Rnd + 1) <span class="vbscomment">'Generates a random number
 from 1-3</span><br>
 Select Case aVar<br>
 &nbsp;&nbsp;&nbsp; Case 1<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alert &quot;You got a one.&quot;<br>
 &nbsp;&nbsp;&nbsp; Case 2<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alert &quot;You got a two.&quot;<br>
 &nbsp;&nbsp;&nbsp; Case Else<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alert &quot;You got something other
 than a one or two.&quot;<br>
 End Select</font></code></p>
</div>
<p><font face="Verdana">Looping structures are used to repeat a section of code,
without having to write it multiple times.&nbsp; Here, two looping structures
will be covered.&nbsp; Together, these two can be used to fill any need for
looping.&nbsp; There are other looping structures which are covered in the <a href="http://www.microsoft.com/Scripting/vbScript/vbslang/vbstoc.htm">MS
VBScript Reference</a>.</font></p>
<div class="vbscode">
 <ol>
  <code>
  <li><font face="Verdana">For <em>indexval</em> = <em>start</em> To <em>finish</em>
   [Step <em>stepval</em>]<br>
   &nbsp;&nbsp;&nbsp; <em>statement(s)</em><br>
   Next</font>
  <li><font face="Verdana">Do While <em>condition</em><br>
   &nbsp;&nbsp;&nbsp; <em>statement(s)</em><br>
   Loop</font></code></li>
 </ol>
</div>
<p><font face="Verdana">The For loop is best used when you know how many times
it should loop.&nbsp; The Indexval is any variable.&nbsp; When the loop is
entered, it becomes equal to Start. &nbsp; Each time it reaches the Next, it
changes the value of Indexval by either Stepval, or one if Stepval is not
specified, until it is past Finish. For loops can be used very efficiently on
arrays.&nbsp; Ex:</font></p>
<div class="vbscode">
 <p><code><font face="Verdana">Dim x<br>
 For x = 3 To 0 Step -1<br>
 &nbsp;&nbsp;&nbsp; If x &gt; 0 Then<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alert x<br>
 &nbsp;&nbsp;&nbsp; Else<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alert &quot;BlastOff!&quot;<br>
 &nbsp;&nbsp;&nbsp; End If<br>
 Next<br>
 Dim y (3)<br>
 For x = 0 To 3<br>
 &nbsp;&nbsp;&nbsp; y(x) = (x + 1) * 3<br>
 &nbsp;&nbsp;&nbsp; Alert y(x)<br>
 Next</font></code></p>
</div>
<p><font face="Verdana">The Do loops (there are more than one) are best used
when you don't know how many times it should loop, such as when reading in from
a file which contains an unknown amount of data.&nbsp; The Do While loop
executes the statements as long as the Condition is True. &nbsp; If, when it
gets to the Do statement, the condition is False, it will not execute the
statements at all.&nbsp; It is important to be careful when making loops,
because if the value of the Condition doesn't change within the loop's
statements, the loop will never stop. Ex:</font></p>
<div class="vbscode">
 <p><code><font face="Verdana">Function RandomString(len)<br>
 &nbsp;&nbsp;&nbsp; Dim lenCount<br>
 &nbsp;&nbsp;&nbsp; Dim tmpString<br>
 &nbsp;&nbsp;&nbsp; Dim tmpChar<br>
 &nbsp;&nbsp;&nbsp; Randomize 'Insure that the numbers are really random<br>
 &nbsp;&nbsp;&nbsp; lenCount = 0<br>
 &nbsp;&nbsp;&nbsp; Do While lenCount &lt; len<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; tmpChar = Chr(Int(75 * Rnd + 48))<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'Random characters (letters,
 numbers, etc.)<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; tmpString = tmpString &amp; tmpChar<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 'Add the character to the string<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; lenCount = lenCount + 1 'Increase
 the counter<br>
 &nbsp;&nbsp;&nbsp; Loop<br>
 &nbsp;&nbsp;&nbsp; RandomString = tmpString 'Return the string<br>
 End Function</font></code></p>
</div>
<p><font face="Verdana">If you're thinking that this could also be done with a
For loop, you're right.&nbsp; In many instances, it is personal preference which
determines which looping structure to use.</font></p>
```

