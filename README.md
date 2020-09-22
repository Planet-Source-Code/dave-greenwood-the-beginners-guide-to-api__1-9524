<div align="center">

## The Beginners Guide To API


</div>

### Description

This article teaches you the basics of Windows API

by giving you a walk through of Declaring API

Functions from Start to Finish & uses real examples

of useful code you can use in your Projects!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Greenwood](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-greenwood.md)
**Level**          |Beginner
**User Rating**    |4.3 (121 globes from 28 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-greenwood-the-beginners-guide-to-api__1-9524/archive/master.zip)





### Source Code

<p><font size="2"><b>The Beginners Guide To API</b></font></p>
<p><font size="2"><b>What is Windows API</b></font></p>
<p><font size="2">It is Windows Application Programming
Interface. This basically means that Windows has built in
functions that programmers can use. These are built into its DLL
files. (Dynamic Link Library)</font></p>
<p><font size="2">So What can these functions do for me (you
might ask) ?</font></p>
<p><font size="2">These pre-built functions allow your program to
do stuff without you actually have to program them.</font></p>
<p><font size="2">Example: You want your VB program to Restart
Windows, instead of your program communicating directly to the
various bits &amp; pieces to restart your computer. All you have
to do is run the pre-built function that Windows has kindly made
for you. This would be what you would type if you have VB4 32, or
higher.</font></p>
<p><font size="2">In your module</font></p>
<p><font color="#000080" size="2"><b>Private</b></font><font
size="2"><b> </b></font><font color="#000080" size="2"><b>Declare</b></font><font
size="2"><b> </b></font><font color="#000080" size="2"><b>Function</b></font><font
size="2"><b> ExitWindowsEx </b></font><font color="#000080"
size="2"><b>Lib</b></font><font size="2"><b> &quot;user32&quot;
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long</b></font></p>
<p><font size="2">If you wanted your computer to shutdown after
you press Command1 then type this In your Form under</font></p>
<p><font size="2">Sub Command1_Click ()</font></p>
<p><font size="2"><b>X = ExitWindowsEx (15, 0) </b></font></p>
<p><font size="2">End Sub </font></p>
<p align="center"><font size="2">----------------</font></p>
<p><font color="#000080" size="2"><b>Private</b></font><font
size="2"><b> </b></font><font color="#000080" size="2"><b>Declare</b></font><font
size="2"><b> </b></font><font color="#000080" size="2"><b>Function</b></font><font
size="2"><b> ExitWindowsEx </b></font><font color="#000080"
size="2"><b>Lib</b></font><font size="2"><b> &quot;user32&quot;
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long</b></font></p>
<p><font size="2">Now to Explain what the above means</font></p>
<p><font color="#000080" size="2"><b>Private</b></font><font
size="2"><b> </b></font><font color="#000080" size="2"><b>Declare</b></font><font
size="2"><b> </b></font><font color="#000080" size="2"><b>Function</b></font><font
size="2"><b> ExitWindowsEx tells VB to Declare a Private Function
called &quot;ExitWindowsEx&quot;. </b></font></p>
<p><font size="2">The<b> </b></font><font color="#000080"
size="2"><b>Lib</b></font><font size="2"><b> &quot;user32&quot; </b>part
tells VB that the function<b> ExitWindowsEx </b>is in the Library<b>
(DLL file) </b>called<b> &quot;user32&quot;. </b></font></p>
<p><font size="2">The final part tells VB to expect the variables
that the<b> ExitWindowsEx </b>function uses<b>. </b></font></p>
<p><font size="2"><b>(ByVal uFlags As Long, ByVal dwReserved As
Long) As Long</b></font></p>
<p><font size="2">The <b>ByVal </b>means pass this variable by
value instead of by reference.</font></p>
<p><font size="2">The <b>As Long </b>tells VB what data type the
information is.</font></p>
<p><font size="2">You can find more about data types in your VB
help files.</font></p>
<p><font size="2">Now you should know what each part of the
Declaration means so now we go on to what does</font></p>
<p><font size="2"><b>X = ExitWindowsEx (15, 0)</b></font></p>
<p><font size="2">For VB to run a function it needs to know where
to put the data it returns from the function. The <b>X = </b>tells
VB to put the response from <b>ExitWindowsEx </b>into the
variable X. </font></p>
<p><font size="2"><b>What's the point of X = </b></font></p>
<p><font size="2">If the function runs or fails it will give you
back a response number so you know what it has done.</font></p>
<p><font size="2">For example if the function fails it might give
you back a number other than 1 so you can write some code to tell
the user this.</font></p>
<p><font size="2">If x &lt;&gt; 1 Then MsgBox &quot;Restart has
Failed&quot;</font></p>
<p align="center"><font size="2">----------</font></p>
<p><font size="2">Now you should know what everything in the
Declaration above means. You are now ready to start using API
calls in your own VB projects. </font></p>
<p><font size="2"><b>To get you started I have included some
useful API calls you might want to use that I've found on Planet
Source Code.</b></font></p>
<p><font size="2"><b>PLAY A WAVEFILE (WAV)</b></font></p>
<p><font color="#000080" size="2">Declare</font><font size="2"> </font><font
color="#000080" size="2">Function</font><font size="2">
sndPlaySound </font><font color="#000080" size="2">Lib</font><font
size="2"> &quot;winmm.dll&quot; </font><font color="#000080"
size="2">Alias</font><font size="2"> &quot;sndPlaySoundA&quot;
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long </font></p>
<p><font color="#000080" size="2">Public</font><font size="2"> </font><font
color="#000080" size="2">Const</font><font size="2"> SND_SYNC =
&amp;H0 </font></p>
<pre>  <font color="#000080">Public</font> <font
color="#000080">Const</font> SND_ASYNC = &amp;H1
  <font color="#000080">Public</font> <font color="#000080">Const</font> SND_NODEFAULT = &amp;H2
  <font color="#000080">Public</font> <font color="#000080">Const</font> SND_MEMORY = &amp;H4
  <font color="#000080">Public</font> <font color="#000080">Const</font> SND_LOOP = &amp;H8
  <font color="#000080">Public</font> <font color="#000080">Const</font> SND_NOSTOP = &amp;H10</pre>
<p><font size="2">Sub Command1_Click ()</font></p>
<p><font size="2">SoundName$ = File 'file you want to play
example &quot;C:\windows\kerchunk.wav&quot; </font></p>
<pre>  wFlags% = SND_ASYNC Or SND_NODEFAULT
  X = sndPlaySound(SoundName$, wFlags%)</pre>
<p><font size="2">End sub</font></p>
<p><font size="2"><b>CHANGE WALLPAPER</b></font></p>
<p><font color="#000080" size="2">Declare</font><font size="2"> </font><font
color="#000080" size="2">Function</font><font size="2">
SystemParametersInfo </font><font color="#000080" size="2">Lib</font><font
size="2"> &quot;user32&quot; </font><font color="#000080"
size="2">Alias</font><font size="2">
&quot;SystemParametersInfoA&quot; (ByVal uAction As Long, ByVal
uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As
Long </font></p>
<pre>
	<font color="#000080">Public</font> <font color="#000080">Const</font> SPI_SETDESKWALLPAPER = 20
<font color="#000080">
</font>Sub Command1_Click ()
<font color="#000080">Dim</font> strBitmapImage As <font
color="#000080">String
</font>strBitmapImage = &quot;c:\windows\straw.bmp&quot;
x = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, 0)</pre>
<p><font size="2">End sub</font></p>
<p><font size="2"><b>ADD FILE TO DOCUMENTS OF WINDOWS MENU BAR</b></font></p>
<p><font color="#000080" size="2">Declare</font><font size="2"> </font><font
color="#000080" size="2">Sub</font><font size="2">
SHAddToRecentDocs </font><font color="#000080" size="2">Lib</font><font
size="2"> &quot;shell32.dll&quot; (ByVal uFlags As Long, ByVal pv
As String)</font></p>
<pre><font color="#000080">Dim</font> NewFile as <font
color="#000080">String
</font>NewFile=&quot;c:\newfile.file&quot;
Call SHAddToRecentDocs(2,NewFile)</pre>
<p><font size="2">MAKE FORM TRANSPARENT</font></p>
<pre><font color="#000080">Declare</font> <font color="#000080">Function</font> SetWindowLong <font
color="#000080">Lib</font> &quot;user32&quot; <font
color="#000080">Alias</font> &quot;SetWindowLongA&quot; _
(ByVal hwnd As Long, ByVal nIndex As Long,ByVal dwNewLong As Long) As Long
<font color="#000080">Public</font> <font color="#000080">Const</font> GWL_EXSTYLE = (-20)
<font color="#000080">Public</font> <font color="#000080">Const</font> WS_EX_TRANSPARENT = &amp;H20&amp;</pre>
<p><font size="2">Private Sub Form_Load()</font></p>
<p><font size="2">SetWindowLong Me.hwnd, GWL_EXSTYLE,
WS_EX_TRANSPARENT</font></p>
<p><font size="2">End</font></p>
<p><font size="2">Any Problems email me at </font><a
href="mailto:DSG@hotbot.com"><font size="2">DSG@hotbot.com</font></a><font
size="2"> </font></p>

