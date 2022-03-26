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
In the Mix Images tab, the app takes two images as input and mixes them to produce a couple of composite images each of which contains the input images overlaid with each other. So it's a *magical operation*! Try it to see what it really means.

In the **Input** section, click the **Browse** buttons to select the first and second input image files that are to be mixed.  
In the **Output** section, click the **Browse** button to select the directory in which the two composite images will be saved. The name of these two composite image files will be in the following format by default:  
*n*.bmp  
where n is an automatically generated number. For example, firstly, the name '1.bmp' is chosen by the app. But if the file '1.bmp' already exists, the name '2.bmp' is chosen; and if it again exists, '3.bmp' is chosen, and so on.

Now start mixing images by clicking the **Mix!** button. Watch the progress bar until the operation is completed, and the following message pops up:

> The images were successfully mixed!

After the mix operation, these two composite image files will be approximately identical to each other. But wait - do not assume that the second file is useless! You are able to *reconstruct* the original input image files from these composite image files if both of them are present. To do so, you need to start a new mix operation with the app by specifying the same composite images as input images. Then the app will create output images that are identical to the original images!

Then what is the **Unit size** field in the Settings section? It's hard to explain, but we recommend that you leave it be the default (1). Because the higher its value, the lower the quality of the composite images.

# thread.htc Component
**thread.htc** is an HTML component (HTC file) that enables the creation of ***virtual threads*** in HTML Applications (HTAs). The term "virtual" means that these threads are not really threads, but are actually *wscript.exe* processes that communicate with the HTA process (mshta.exe) via COM. Moreover, the code that is executed by these threads is originally stored within the HTA file. And as soon as the thread starts, the code is dynamically transfered to the 'wscript.exe' process for execution.

To use the component in an HTA, the first step is to declare an XML namespace named 't', and then import 'thread.htc' into that namespace [(See 'Importing a Custom Element' section in 'About Element Behaviors')](https://docs.microsoft.com/en-us/previous-versions//ms531426(v=vs.85)?redirectedfrom=MSDN).

Then you must define one or more thread templates, which contain the JScript code to be executed by a thread. Then you can create from the thread template as many threads as desired, all of which execute the code within the template. Use the `<t:thread>` element to define a thread template. This element must necessarily have an id. For example, the following piece of code defines a template for a thread that simply moves a large file from one location to another, and then displays a message to indicate the success of the operation:

    <t:thread id="fileMoverThread">
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    fso.MoveFile("C:\\movie.mp4", "D:\\movie.mp4");
    window.document.body.innerText = "The file movie.mp4 has been moved to drive D.";
    </t:thread>

**Note:** The id assigned to the `<t:thread>` element (in this example, `fileMoverThread`) is called the *Thread Template Identifier (TTID)* of the potential threads.

**Note:** Currently, thread templates only support **JScript code**, not code in other languages such as VBScript or PerlScript.

Now, the next step is to create a thread from the thread template, which causes the code within our `fileMoverThread` to start running concurrently with the usual scripts in the HTA. To do so, just call the `start` method of the `<t:thread>` element via a usual script. For example, the following code snippet displays a button that - when clicked - creates a thread to move the video file.

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

then the thread would look for the `tmid` variable in the usual script of the HTA; and it would find its main identifier and display it.

In addition, in thread templates, you must not omit the name of the window object when calling one of its methods or accessing one of its properties. So the following code would also fail:

    <t:thread id="fileMoverThread">
    alert("The thread identifier is " + window.tmid);
    ...

## Problem with viewing thread templates in text editors
One issue with `<t:thread>` elements might be, although they work perfectly, they don't render perfectly in text editors, like Notepad Plus Plus. This is basically because these text editors do not understand the `<t:thread>` element contains JScript code and not plain text; so they don't format and colorize the code properly!

Fortunately, there is a solution: just enclose the code within `<script>` and `</script>` tags. In simpler words, write

    <t:thread id="foo">
    <script>
    // The actual code
    </script>
    </t:thread>

instead of

    <t:thread id="foo">
    // The actual code
    </t:thread>

This way, the JScript code within `<t:thread>` element renders correctly in text editors, and the *thread.htc* component smartly detects and removes the extra `<script>` tags before passing the code to the script engine for execution.
