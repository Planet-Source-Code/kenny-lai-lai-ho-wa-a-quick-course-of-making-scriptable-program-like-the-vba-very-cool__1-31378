<div align="center">

## A quick course of making scriptable program, Like the VBA \(Very Cool\!\)

<img src="PIC200221923293962.gif">
</div>

### Description

<p align="center"><b><font face="Arial" color="#000080">Scriptable make

everything possible possible</font></b></p>

<p>Have you ever use the VBA in Microsoft Office? Making your application

scriptable can enable it's functions to be extent to infinite, by the End Users.

End Users can &quot;WRITE PROGRAM ON YOUR PROGRAM&quot;, and run it as they

like. It sounds interesting?</p>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kenny Lai, Lai Ho Wa](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kenny-lai-lai-ho-wa.md)
**Level**          |Advanced
**User Rating**    |4.9 (69 globes from 14 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kenny-lai-lai-ho-wa-a-quick-course-of-making-scriptable-program-like-the-vba-very-cool__1-31378/archive/master.zip)





### Source Code

<p>A quick course of making <font size="4"><b>scriptable</b></font> program,
Like the VBA <b>(Very Cool!)</b></p>
<p align="center"><b><font face="Arial" color="#000080">Scriptable make
everything possible possible</font></b></p>
<p>Have you ever use the VBA in Microsoft Office? Making your application
scriptable can enable it's functions to be extent to infinite, by the End Users.
End Users can &quot;WRITE PROGRAM ON YOUR PROGRAM&quot;, and run it as they
like. It sounds interesting?</p>
<p>@</p>
<p>This is a quick course teaching you how to make you application scriptable,
using Microsoft Scripting Control.</p>
<hr>
<p align="center"><b><font face="Arial" color="#000080">Understanding Microsoft
Scripting Control</font></b></p>
<p>This is a free gift come together with Visual Basic. It support VBScript and
JScript. But for convinence, I will use VBScript for demonstration.&nbsp;</p>
<p>It is very easy to use. Let's say we have a script control SC</p>
<p><font color="#000080">Private</font> <font color="#000080"> Sub</font> Command1_Click()<br>
<br>
<font color="#000080">&nbsp;&nbsp;&nbsp;</font> <font color="#000080">Dim</font> strProgram
<font color="#000080"> As</font> String<br>
</p>
<blockquote>
 <p>	strProgram = "Sub Main" &amp; vbCrLf &amp; _<br>
	"MsgBox ""Hello World&quot;&quot;&quot; &amp; vbCrLf &amp; _<br>
	"End Sub"<br>
 <br>
	sc.language = "VBScript"</p>
 <p><br>
	sc.addcode strProgram<br>
	sc.run "Main"</p>
</blockquote>
<p><br>
<font color="#000080">End Sub</font></p>
<p>A message box will appear when you press Command1. The code is in VBScript
format(*) and can be enter by any method you like, said TextBox. This enable
end-users entering their own VBScript code they like, and run them. It just like
another Visual Basic!</p>
<p>(* The main difference is that the only varible type is viarant. e.g. Dim
a,b,c but NOT Dim a as string)</p>
<p>So, what can you do to make your application scriptable, extentable?</p>
<hr>
<p align="center"><b><font face="Arial" color="#000080">Program Overview</font></b></p>
<p>Now, right click on the controls list and add a reference to &quot;Microsoft
Script Control 1.0&quot;. Create one on a form.</p>
<table border="1" cellpadding="0" cellspacing="0" width="100%">
 <tr>
  <td width="50%" align="center"><font color="#000080">Name</font></td>
  <td width="50%" align="center"><font color="#000080">Type</font></td>
 </tr>
 <tr>
  <td width="50%">SC</td>
  <td width="50%">Microsoft Script Control</td>
 </tr>
 <tr>
  <td width="50%">Form1</td>
  <td width="50%">Form</td>
 </tr>
 <tr>
  <td width="50%">Text1</td>
  <td width="50%">TextBox1</td>
 </tr>
 <tr>
  <td width="50%">txtCode</td>
  <td width="50%">TextBox</td>
 </tr>
 <tr>
  <td width="50%">txtCommand</td>
  <td width="50%">TextBox</td>
 </tr>
 <tr>
  <td width="50%">lstProcedures</td>
  <td width="50%">ListBox</td>
 </tr>
 <tr>
  <td width="50%">CmdRun</td>
  <td width="50%">Command Button</td>
 </tr>
</table>
<p>@</p>
<p>The Text1 is used as an object that is the &quot;Scriptable&quot; part. In
this program, end users can enter Visual Bascis SCRIPT code in txtCode. They may
run the code by entering command lines in txtCommand, and press CmdRun.</p>
<hr>
<p align="center"><font color="#000080" face="Arial"><b>The main part</b></font></p>
<p>There is a AddObject function in the script control. You can add any object,
controls, like textbox, forms, buttons and picture box into the script control,
and give them a &quot;scripting name&quot;, i.e. the name used to identify the
object in the end-users code.</p>
<p><font color="#000080">Private</font> Form_Load</p>
<p>sc.AddObject &quot;MyText&quot;, Text1</p>
<p><font color="#000080">End Sub</font></p>
<p>After add the Text1 into the script control, you can access the Text1 in the
End-users code that is entered in the txtCode.</p>
<p>e.g.</p>
<p>In the txtCode, enter the following code:</p>
<p><i><b>Sub Main</b></i></p>
<p><i><b>&nbsp;&nbsp;&nbsp; Msgbox MyText.Text</b></i></p>
<p><i><b>End Sub</b></i></p>
<p>Also add the following code to the program(Not the textbox)</p>
<p><font color="#000080">Private Sub</font> CmdRun_Click</p>
<p>&nbsp;&nbsp;&nbsp; sc.run &quot;Main&quot;</p>
<p><font color="#000080">End Sub</font></p>
<p><font color="#000080">Private Sub</font> sc_Error()</p>
<p>&nbsp;&nbsp;&nbsp; MsgBox "Error running code: " &amp; SC.Error.Description &amp; vbCrLf &amp; "Line:" &amp;
SC.Error.Line</p>
<p><font color="#000080">End Sub</font></p>
<p>When you click the CmdRun, the code in the txtCode Sub Main section will be
run. Now you can see the how easy to control the program by end-uses code. The
&quot;Msgbox ...&quot; can be replace by any logical VBScript code. E.g. if you
entered MyText.visible=False, the textbox will disappear.</p>
<p>Similarly, you can AddObject of any controls and object you like into script
control and control it totally by end-users code. This is the basis of making
scriptable application.</p>
<p>Futhermore, the script control provide the procedures object so that you can
get all information of the procedures of your code.</p>
<p><font color="#000080">Private Sub</font> txtCode_Change</p>
<p><font color="#000080">&nbsp;&nbsp;&nbsp; On Error Resume Next</font></p>
<p>&nbsp;&nbsp;&nbsp; lstProcedures.Clear</p>
<p>&nbsp;&nbsp;&nbsp; <font color="#000080">Dim</font> i <font color="#000080">as</font>
integer</p>
<p>&nbsp;&nbsp;&nbsp; <font color="#000080">For</font> i=1 <font color="#000080">to</font>
sc.Procedures.Count</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; lstProcedures.Additem
sc.Procedures(i)</p>
<p>&nbsp;&nbsp;&nbsp; <font color="#000080">Next</font> i</p>
<p><font color="#000080">End Sub</font></p>
<p><font color="#000080">Sub</font> ExecuteCommand(Str <font color="#000080">As</font>
string)</p>
<p>&nbsp;&nbsp;&nbsp;<font color="#000080"> On Error Goto</font> 1</p>
<p>&nbsp;&nbsp;&nbsp; sc.ExecuteStatement Str</p>
<p><font color="#000080">Exit Sub</font></p>
<p>1</p>
<p>Msgbox Error</p>
<p><font color="#000080">End Sub</font></p>
<p>For the ExecuteCommand, your can enter a correct statement to execute like:</p>
<p>Msgbox MyText.Text</p>
<p>Main</p>
<p>MyProcedures Arg1,Arg2</p>
<p>Msgbox MyFunction(Arg1, Arg2, Arg3)</p>
<hr>
<p align="center"><font color="#000080" face="Arial"><b>Demonstration Program</b></font></p>
<p>And by now, you may be able to create your scriptable program, or make a
scripting console for your application. </p>
<p>Here is my demonstration program, my Page Creator 3. I don't like to write
just a simple program. </p>
<p>On the Left hand Outlook bar, click on the PCScript Console to open the
console panel. Open a template by clicking open. After entering your own code
you like, return the main program or the HTML Code Editor. Find the tab
&quot;PCScript&quot; on the floating toolbox. There is a command window. Just
type the command in the textbox and click return key to execute it. Have Fun!</p>
<p>Kenny Lai</p>
<p>Download Demonstration:</p>
<p><a href="http://student.mst.edu.hk/~s9710050/Page%20Creator%203.zip">http://student.mst.edu.hk/~s9710050/Page
Creator 3.zip</a></p>
<p>@</p>

