# Threading
An HTML component which enables you to create threads in HTAs, along with an example of multi-threaded app that makes use of this component.

# Example app
Run the app by double-clicking the file 'multithreading example.hta'. The app has two tabs: **Search** and **Mix Images**.

## Search tab
In the Search tab, you can search your file system. In the **Location** field, enter the path of the directory in which the app will look for files. In the **File name** field, enter the name of the file you wish to find. Note that you can also place *wildcards* in the file name.

Finally, click the **Search** button to start the search. Technically, this will result in creation of a thread that performs the search operation. After the search is started, you can stop it by simply clicking the **Stop** button. Technically, this will result in the search thread being killed.

The results of the search will be displayed in the light-yellow box below the buttons. Clicking on each result will make the app open the containing folder of the file in a File Explorer window. You can also clear the list of results by clicking the **Clear Results** button next to the Search button.

Clicking **Read Results** button will make the app create a thread to read the search results aloud via the system's text-to-speech engine. After clicking it, you can stop the playback by clicking it again.

## Mix Images tab
**Note:** To use this tab, you need to run the app on **Windows Vista or later**.

In the Mix Images tab, the app takes two images as input and mixes them to produce a couple of composite images each of which contains the input images overlaid with each other. So it's a *magical operation*! Try it to see what it really means.

In the **Input** section, click the **Browse** buttons to select the first and second input image files that are to be mixed.  
In the **Output** section, click the **Browse** button to select the directory in which the two composite images will be saved. The name of these two composite image files will be in the following format by default:  
*n*.bmp  
where n is an automatically generated number. For example, firstly, the name '1.bmp' is chosen by the app. But if the file '1.bmp' already exists, the name '2.bmp' is chosen; and if it again exists, '3.bmp' is chosen, and so on.

Now start mixing images by clicking the **Mix!** button. Watch the progress bar until the operation is completed, and the following message pops up:

> The images were successfully mixed!

After the mix operation, these two composite image files will be approximately identical to each other. But wait - do not assume that the second file is useless! You are able to *reconstruct* the original input image files from these composite image files if both of them are present. To do so, you need to start a new mix operation with the app by specifying the same composite images as input images. Then the app will create output images that are identical to the original images!

Then what is the **Unit size** field in the Settings section? It's hard to explain, but we recommend that you leave it be the default (1). Because the higher its value, the lower the quality of the composite images.

If the **Preserve image transparency** option in the Settings section is checked, then if any of the pixels in the input images are transparent, the corresponding pixels in the composite images will be also transparent. Otherwise, the app will generate all the pixels in the composite images with an alpha value of 255, so that the entire composite images will be opaque.

# Omegathread.htc Component
[HTML Applications (HTAs)](https://en.wikipedia.org/wiki/HTML_Application) is a great technology that provides a way to write ordinary Microsoft Windows programs using Dynamic HTML and scripting languages (e.g. JScript or VBScript), enabling these languages to run outside of conventional web browsers like Chrome or Firefox. These applications, despite all their benefits, are known to be **single-threaded**, which might cause some problems when developing processor-intensive applications. For example, imagine you are going to make an HTA that is supposed to move a large file (e.g. a 4-GB file) from one given location to another; so you incorporate a form in your HTA that contains two text fields which take the source and destination file paths from the user. In the form, there is also a button which, when clicked, will perform the file movement operation by calling the [`MoveFile`](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/movefile-method) method of the [`FileSystemObject`](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object). Since the `MoveFile` method does a synchronous operation and the given source file is very large, the script of the HTA blocks at `fso.MoveFile( ... )` command for a long time. This will unfortunately result in the HTA window and its whole user interface freezing up (=hanging) until the file movement operation is completed. Thus, the end-user will be probably unhappy with your app.

The **Omegathread library** has been designed to solve that problem by providing multithreading capabilities for HTAs. With the aid of this library, in the preceding example, clicking the form's submit button will cause the program to create a thread which performs the file movement operation, and when the operation is completed, the thread will display a message to the user, indicating that the file has been successfully moved. This way the user interface won't freeze up, and the end-user won't be frustrated.

**Omegathread.htc** is an HTML component (HTC file) that enables the creation of ***virtual threads*** in HTML Applications (HTAs). The term "virtual" means that these threads are not really threads, but are actually *wscript.exe* processes that communicate with the HTA process (mshta.exe) via COM. Moreover, the code that is executed by these threads is originally stored within the HTA file. And as soon as the thread starts, the code is dynamically transfered to the 'wscript.exe' process for execution.

## Basic usage
To use the component in an HTA, first, copy the two files "omegathread.htc" and "thread_host.wsf" (other ones are not required) from this repository to the directory where your HTA is stored. Then follow these steps:

The first step is to declare a namespace named 't', and then import the component into that namespace. The namespace 't' can be declared by setting the `xmlns` attribute in the `<html>` element (the root element of your HTA), as follows:

    <html xmlns:t>

Then the 'omegathread.htc' component must be imported into the namespace 't' by adding the following `<?import?>` directive to the `<head>` element.

    <?import namespace="t" implementation="omegathread.htc" ?>

The next step is to define one or more thread templates, which contain the script code to be executed by a thread. This way you can later create from the thread template as many threads as desired, all of which execute the code within the template. Use the `<t:thread>` element to define a thread template. This element must necessarily have an id. For example, the following piece of code defines a template for a thread that simply moves a large file from one location to another, and then displays a message to indicate the success of the operation:

    <t:thread id="fileMoverThread">
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    fso.MoveFile("C:\\movie.mp4", "D:\\movie.mp4");
    window.document.body.innerText = "The file movie.mp4 has been moved to drive D.";
    </t:thread>

**Note:** The id assigned to the `<t:thread>` element (in this example, `fileMoverThread`) is called the *Thread Template Identifier (TTID)* of the potential threads.

**Note:** In the preceding example, the code within the thread template is written in JScript. If you need to place VBScript code in your thread template, you must specify `vbscript` as the `slanguage` attribute of the `<t:thread>` element. For example, the following is the VBScript equivalent of the `fileMoverThread`:

    <t:thread id="fileMoverThread" slanguage="vbscript">
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile "C:\movie.mp4", "D:\movie.mp4"
    window.document.body.innerText = "The file movie.mp4 has been moved to drive D."
    </t:thread>

Now, the next step is to create a thread from the thread template, which causes the code within our `fileMoverThread` to start running. To do so, just call the `start` method of the `<t:thread>` element via a usual script. For example, the following code snippet displays a button that - when clicked - creates a thread to move the video file.

    <script language="jscript">
    var tmid;
    function moveFileByThread() {
        tmid = fileMoverThread.start();
    }
    </script>
    ...
    <button onclick="moveFileByThread()">Move video file</button>

As you see in the code, the `start` method returns a value that is assigned to the `tmid` variable. This value is called **Thread Main Identifier (TMID)**, which is a numeric value that uniquely identifies the thread among all the threads that are created from the template. It can be used in subsequent calls to methods such as `terminate`, `getExitCode`, .etc to terminate the thread, retrieve its exit code, .etc. It contrasts to the previously-mentioned TTID, which was used to identify the `<t:thread>` element (the thread template) throughout the HTML document.

## Passing a parameter to a thread
What's interesting is that you can pass any desired value (like a number, a string, an object, .etc) to the `start` method. Then this value can be accessed by the thread through a special variable named `tparam`. For example, the following code snippet defines a template for a thread that deletes a file.

    <t:thread id="deleterThread">
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    fso.DeleteFile(tparam);
    </t:thread>

The pre-defined variable `tparam` is passed to the `DeleteFile` method. Here, the presence of `tparam` indicates that we don't directly specify the path of the file that should be deleted. Instead, when we create the thread, we send it the path of the file; like below:

    var tmid = deleterThread.start("C:\\Users\\Bob\\Desktop\\note.txt");

Now we want to extend our previous example by letting the user select their desired source and destination file names to move - via a form. Here is the full source code:

    <html xmlns:t>
    <head>
    <title>Move files</title>
    <?import namespace="t" implementation="thread.htc" ?>

    <t:thread id="fileMoverThread">
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    fso.MoveFile(tparam.source, tparam.dest);
    window.document.all.moveButton.innerText = "Move";
    </t:thread>

    <script language="jscript">
    var tmid;

    function moveFileByThread() {
        tmid = fileMoverThread.start({
            source: document.all.fileSource.value,
            dest: document.all.txtDest.value
        });
        
        document.all.moveButton.innerText = "Moving...";
    }
    </script>
    </head>
    <body>
    <form>
    <label for="fileSource">Source file:</label>
    <input type="file" id="fileSource"><br>
    <label for="txtDest">Destination file:</label>
    <input type="text" id="txtDest"><br>
    <button id="moveButton" onclick="moveFileByThread()">Move</button>
    </form>
    </body>
    </html>

## Scope of variables and functions
Please note that the variables defined within the thread template, and also the pre-defined `tparam` variable are all local to the threads of the template. As with our File Mover example, the usual script can't access the `fso` variable that is defined within the thread template. The same applies to functions.

However, the code in the thread template can access all of the global variables defined in the usual script, provided that it prepends `window.` to the variable name. For example, imagine the `fileMoverThread` element contained the code below:

    <t:thread id="fileMoverThread">
    window.alert("The thread identifier is " + tmid);
    ...

Then the thread would generate a run-time error, because there is no `tmid` variable defined in the template. But if the code were

    <t:thread id="fileMoverThread">
    window.alert("The thread identifier is " + window.tmid);
    ...

then the thread would look for the `tmid` variable in the usual script of the HTA; and it would determine its main identifier and display it.

### Using the `extern` function
Consider this command is used in a thread template:

    var foo = window.foo;

This implies that a variable named `foo` was defined in the context of the usual script, and now we want to copy its value into a variable with the same name in the context of the thread template. This is a useful technique if you need to reference such a variable many times in the thread template and you don't want to repeatedly prepend `window.` to the variable name.

But doing so is even easier if you call the `extern` function, which takes a variable name as a string parameter and copies the variable from the context of the HTA script to the context of the thread. For example, the following command has the identical functionality as the preceding example:

    extern("foo");

In the example below, multiple variable names are passed to `extern`, so that all of those variables will be copied to the context of the thread.

    extern("x", "y", "z");

By the way, in thread templates, you must not omit the name of the window object when calling one of its methods or accessing one of its properties. So the following code would also fail:

    <t:thread id="fileMoverThread">
    alert("The thread identifier is " + window.tmid);
    ...

## Problem with viewing thread templates in text editors
One issue with `<t:thread>` elements might be, although they work perfectly, they don't render perfectly in text editors, like Notepad Plus Plus. This is basically because these text editors do not understand the `<t:thread>` element contains script code and not plain text; so they don't format and colorize the code properly!

For JScript-based thread templates, there is a solution: just enclose the code within `<script>` and `</script>` tags. In simpler words, write

    <t:thread id="foo">
    <script>
    // The actual code
    </script>
    </t:thread>

instead of

    <t:thread id="foo">
    // The actual code
    </t:thread>

This way, the JScript code within `<t:thread>` element renders correctly in text editors, and the *omegathread.htc* component smartly detects and removes the extra `<script>` tags before passing the code to the script engine for execution.

## Storing the thread code in external files
Instead of placing the thread code directly within the `<t:thread>` element, you can use the `src` attribute of the thread template element to instruct the system to load the specified file and obtain the thread code. It's so simple; consider we have a script file named "HelloWorld.vbs" with the code below:

    MsgBox "Hello, World!"

Then we define a thread template in an HTA like below:

    <t:thread id="helloThread" src="HelloWorld.vbs" />

Finally, the following command can be used in a script to create a thread which simply displays the message, "Hello, World!".

    helloThread.start();

## Exitting from a thread
Sometimes you need to exit from a thread and make it avoid executing the rest of its code; like the Win32 `ExitThread` function does in C++. Then you're welcome to use the following command in the thread code:

    threadc.exit();

You can also pass an exit code to the `exit` method, as follows:

    threadc.exit(100);

## Pausing a thread for a particular period
A thread can call the `sleep` method of the `threadc` object to pause its execution for a specified period of time:

    threadc.sleep(number of milliseconds);

For example the following command pauses the execution of a thread for 4 seconds:

    threadc.sleep(4000);

## Rendezvous Port
Unfortunately, apps that use `omegathread.htc` component create an extra window that sticks on the taskbar, but displays no content or user interface. This window is called **"The Rendezvous Port"** and must not be closed by the user since all threads require it to connect to the HTA process. Once this window is closed, whenever the app attempts to create a new thread, an error message similar to the following picture pops up:

![rendezvous port closed](https://user-images.githubusercontent.com/31417320/161370903-fcbd1be5-352d-4625-90f0-a08741b6f887.png)

Moreover, the `omegathread.htc` component automatically appends to the `<body>` element an `<object>` element which is related to the Rendezvous Port. This element doesn't display any additional content in your HTA and must not be removed from the document tree.

## Debugging threads
Whenever you simply create a thread (e.g. by calling `myThread.start()`), active debugging for that thread is disabled by default. For example, using the JScript `debugger` statement in the thread code will cause nothing to happen. Additionally, you are generally unable to use any debugger program to attach to the thread and set breakpoints in it. Attempting to do so with Microsoft Visual Stutio, for example, will cause the debugger to keep showing the following message:

> Waiting to break when the next script code runs...

So if you want to debug your thread, first, set the `debugmode` attribute of the thread template element to `enabled`, like below:

    <t:thread id="fileMoverThread" debugmode="enabled">

Now you can either use a `debugger` statement in the thread template code, or do the following steps in order to attach your debugger app to the thread and set breakpoints in it.

Launch your HTA. Then launch your debugger app and attach to the HTA process. Next, make your HTA create the thread that you wish to debug. Upon thread creation, the OmegaThread component logs a message in the debugger's output window which is similar to the following:

    # OmegaThread: A thread was created from the template 'fileMoverThread'.
    # TMID: 1, THID: 2435

In this message, the number 1 indicates the **Thread Main Identifier (TMID)**, and 2435 indicates the **Thread Host Identifier (THID)**. So now you must open a new debugger window, and attach to a process named `wscript.exe` whose process ID is the same as THID (in this example, 2435). This process is called, "the **Thread Host Process**".

Then imediately click the "Break All" button in the debugger. So now the debugger breaks the thread, displays its code, and lets you debug it.

Alternatively, you can set the `debugmode` attribute to `auto`. In this case, you won't have to manually perform the above steps in order to attach the debugger to the thread. Rather, as soon as the thread starts, the debugger automatically launches and breaks the thread at the beginning of its code. The following example shows how to set `debugmode` attribute to `auto`:

    <t:thread id="fileMoverThread" debugmode="auto">

The following information applies to older versions of OmegaThread. It describes another debugging approach that still works, but is no longer recommended.

~~To prevent this situation, when calling the `start` method to create a thread, you must mark the thread as debugable by setting the third parameter of the `start` method to `true`. Please note that this method has also a second parameter which we are not going to cover in this document; so you can just simply set it to `false`. The following code snippet starts the `fileMoverThread` from the previous examples with active debugging enabled.~~

    tmid = fileMoverThread.start({
        source: document.all.fileSource.value,
        dest: document.all.txtDest.value
    }, false, true);

~~Now you can debug your thread using any desired script debugger app. Please note that each of the threads created by the *Omegathread.htc* component actually runs in the context of a special process named **`wscript.exe`**. So all you need to do is open the list of processes in your debugger app, find the `wscript.exe` process, and attach to it. Alternatively, if you are using JScript for your thread code, you can place a `debugger` statement anywhere in the thread code, so that when the thread execution reaches this statement, the debugger app automatically launches and attaches to the thread.~~

## Use of OmegaThread in image editing
The [WIA Automation Layer](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-startpage) enables scripts to edit raster image files (e.g. BMP, JPEG, PNG, GIF, TIFF files) by allowing scripts to have control over all the pixels of the image, and determine/modify the color of each pixel. For this purpose, the **WIA Automation** provides a set of COM objects including the [`ImageFile`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-imagefile) and [`Vector`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-vector) objects. After a script instantiates the [`ImageFile`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-imagefile) object, it calls its [`LoadFile`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-iimagefile-loadfile) method to load the raster image file of interest. The script then invokes the [`ARGBData`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-iimagefile-argbdata) property of the [`ImageFile`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-imagefile) object, which returns a [`Vector`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-vector) object containing the colors of all the pixels of the image. Finally, in order to determine/modify the color of a particular pixel, the script invokes the [`Item`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-ivector-item) property of the [`Vector`](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-vector) object, passing it the index of the pixel of interest. The **`Item`** property sets or retrieves a **Long** value representing the **Alpha, Red, Green, and Blue** color components of the pixel. So WIA Automation is very useful for scripts to edit raster image files.

On the other hand, using the **WIA Automation** to perform a heavy operation on an entire raster image may have a bad impact on the performance of the script, which is a common problem that is also discussed [on this page](https://www.vbforums.com/showthread.php?889525-WIA-2-0-Vector-Add()-method-slow). For example, suppose you want to write a script that converts a raster image file to grayscale. Then you would have to iterate through all the pixels of the image, in order to calculate the average of the RGB component values of each pixel, and then change the value of that pixel accordingly. Since there is *an abundance of pixels* in the image, iterating through *all the pixels* in the image would cause the operation to take a lot of time, and would result in a very slow performance.

That's where the **OmegaThread library** comes to help. The problem can be solved by writing a multi-threaded HTA and having the conversion operation performed by a separate thread; as a result, the user interface of the HTA remains responsive throughout the operation, so that the HTA can display a progress bar or a loading spinner to evoke patience in the end-user.

For instance, below is a **multi-threaded** HTA that converts a raster image file to grayscale. In the HTA, there is a form with an `<input type="file">` element which allows the user to select the raster image file of interest. Within the form, there is also a button named **"Gray-Scale"** that the user can click to start the conversion operation. Internally, when clicking the **Gray-Scale** button, the HTA creates a separate worker thread (via the **OmegaThread library**) which performs the operation. Then the HTA displays [a loading spinner](https://javadbayat.github.io/Test/processing.gif) (as a gif image) and keeps it visible until that thread finishes its work. When the conversion operation is over, the worker thread displays a `prompt` dialog box which asks the user to enter the path of the file in which the resulting grayscale image is to be stored.

```
<html xmlns:t>
<head>
<title>Make Gray-Scale Photos</title>
<hta:application windowstate="maximize" scroll="no" contextmenu="no" />
<meta http-equiv="content-type" content="text/html; charset = UTF-8">
<meta http-equiv="MSThemeCompatible" content="Yes">

<?import namespace="t" implementation="omegathread.htc" ?>

<script language="jscript">
var tmid = 0;

onload = function() {
    document.mainForm.btnMakeGrayscale.onclick = startProcessing;
};

function startProcessing() {
    tmid = grayscaleThread.start({
        imageFilePath: document.mainForm.imageFile.value,
        indicateCompletion: function() {
            document.mainForm.imageFile.disabled = false;
            document.mainForm.btnMakeGrayscale.innerText = "Gray-Scale";
            document.mainForm.btnMakeGrayscale.disabled = false;
            document.images.waitSign.style.display = "none";
            document.body.style.cursor = "default";
        }
    });
    
    document.mainForm.imageFile.disabled = true;
    document.mainForm.btnMakeGrayscale.innerText = "Processing";
    document.mainForm.btnMakeGrayscale.disabled = true;
    document.images.waitSign.style.display = "inline";
    document.body.style.cursor = "wait";
}
</script>

<t:thread id="grayscaleThread" slanguage="jscript">
<script>
var fso = new ActiveXObject("Scripting.FileSystemObject");
if (!fso.FileExists(tparam.imageFilePath)) {
    window.alert("The specified image file was not found.");
    threadc.exit();
}

var wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";

var imageFile = new ActiveXObject("WIA.ImageFile");
imageFile.LoadFile(tparam.imageFilePath);

var nPixels = imageFile.Width * imageFile.Height;
var vector = imageFile.ARGBData;
var mask = 0xFF << 24;
for (var i = 1; i <= nPixels; i++) {
    var color = vector(i);
    var red = color & 0x00FF0000 >> 16;
    var green = color & 0x0000FF00 >> 8;
    var blue = color & 0x000000FF;
    var average = Math.floor((red + green + blue) / 3);
    color = color & mask | (average << 16) | (average << 8) | average;
    vector(i) = color;
}

var suggestion = getFileNameSuggestion();
var outputImageFile = window.prompt("Enter the full name of the file in which to store the output image.", suggestion);
if (!outputImageFile)
    threadc.exit();

var ip = new ActiveXObject("WIA.ImageProcess");
ip.Filters.Add(ip.FilterInfos("ARGB").FilterID);
ip.Filters(1).Properties("ARGBData") = vector;
ip.Filters.Add(ip.FilterInfos("Convert").FilterID);
ip.Filters(2).Properties("FormatID") = wiaFormatPNG;

imageFile = ip.Apply(imageFile);
imageFile.SaveFile(outputImageFile);

tparam.indicateCompletion();

function getFileNameSuggestion() {
    var filePath = tparam.imageFilePath;
    var newFileName = "grayscale_" + (fso.GetBaseName(filePath)) + ".png";
    return fso.BuildPath(fso.GetParentFolderName(filePath), newFileName);
}
</script>
</t:thread>
</head>
<body>
<form name="mainForm" onsubmit="return false;">
    <label for="inputImageFile">Image File:</label>
    <input type="file" name="imageFile" id="inputImageFile"><br><br>
    <button name="btnMakeGrayscale">Gray-Scale</button>
</form>
<img name="waitSign" src="https://javadbayat.github.io/Test/processing.gif" style="display: none;">
</body>
</html>
```

**Code explanation:**  
The function **`startProcessing`** handles the **`onclick`** event for the **Gray-Scale** button. This function takes the following actions:

+ Creates **the worker thread** to perform the conversion operation. Also passes the thread a parameter which will be explained later.
+ Changes the text within the **GrayScale** button from `Gray-Scale` to `Processing`. Then disables all the elements within the form (because the program can handle only one conversion operation at a time, and the user must not enter anything in the form while the operation is ongoing.)
+ Makes the loading spinner visible by changing the value of CSS attribute `display` from `none` to `inline`.
+ Changes the cursor of the program in order to indicate the system is working.

When **the worker thread** is created, it is passed **an object parameter** that contains the following properties:

+ **`imageFilePath`**: The path of the raster image file that is to be converted. This path is taken from the value of the `<input type="file">` element within the form, and then it is assigned to this property.
+ **`indicateCompletion`**: A pointer to a call-back function which is called by the worker thread after the conversion operation is completed. This call-back function merely restores the HTA user interface to its original state (e.g. makes the loading spinner hidden .etc).

When the worker thread starts to run, it first does some validation. So it checks the existance of the input image file that was specified on the form. If the file does not exist, then the worker thread calls `window.alert` to display an error message and then exits.

Next, the worker thread starts utilizing the **WIA Automation**. It loads the input image file, and modifies its pixels to make them grayscale. Then the thread calls `window.prompt` to open a dialog box to ask the user to enter the path of the output image file. Finally, it stores the output image file to disk (in PNG format).
