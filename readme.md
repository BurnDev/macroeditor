# Macro Editor Help Guide
Excel formulas pack a powerful punch, but sometimes you need a little more. You need functionality specific to your business rules or productivity requirements, something that is just not possible with predetermined methods and existing functions. Excel attempts to address this gap with macros, allowing the user to create their own automated tasks with one (BIG) caveat...macros are coded in VBA.

But with less than 5% popularity in the industry, VBA is quite an impediment to actually getting the job done. Enter _**Macro Editor**_.

## The Projects Goals
Our motto is 'Your macro, your language'. If you need additional Excel functionality, go ahead and do it your way. Write your own code, in any technology of your choice, and your macro will just work. Simple as that. 

Currently we support both JavaScript and Python fan but we are building out support for all popular languages. Our goal is to bring macros to the masses.

## The Demo
[Macro Editor](https://youtu.be/QY2bL1wtXvI)

## The Basics
![](/images/macro-editor.jpg "Macro Editor")

### The Editor Section
This is the main area of the app and is where you will be writing all your awesome scripts. Pretty self explanatory (we hope)!

### The Terminal Section
Below the main coding area, we have the **Terminal Section**.

The **Terminal Section** has two tabs, *OUTPUT* and *EXAMPLES*
* The *OUTPUT* tab captures and displays the log stream from your scripts. Any time you do a ```console.log``` or a ```print``` or ```cout<<```, (etc, etc) this is where you will be checking the output. 

    See **How To Get Your Results To Interact With Excel** below for more information on how to do this.
* The *EXAMPLES* tab is where you can see a list of pre built examples in multiple languages to help you get started. Clicking on the example link will load the file in the editor, and it will run after you hit execute. Typically, each example comes in a set of two scripts, one to help simulate the data, and another to validate or manipulate that data. We have a great explanation of this in our demo, where we review our Tic Tac Toe example in detail.

To the right of the tabs you will see the *up* and *down* arrows.  The **Terminal Section** is a split section and can be dragged to resize the area.  The *up* arrow will maximize the **Terminal Section**, and the *down* arrow will minimize it.

To the right of the *up* and *down* arrows you will see the *CLEAR* button.  This will clear all logged outputs within the *OUTPUT* tab.

### Language Selection
In the lower left hand corner you will see the currently selected language. To change this, click on the language and you will see a dropdown of our supported languages; currently JavaScript and Python with more in development.

> Note: We plan on adding a lot more languages in the near future, all of which will get added to this dropdown as we add support - so keep checking back for more!

### Cell Details
To the right of the **Language Selection** button you will see the **Cell Details** checkbox. This tells the editor the kind of information you want from excel for a particular cell.  
* Without **Cell Details** enabled, data from Excel will be imported as an array of values.
* With **Cell Details** enabled, data from Excel will be imported as an array of objects elements where each object contains the *cell location*, *column letter*, *row number*, and *value* of the cell.
> For a more detailed explanation of how cell details works read through the **How To Get Data From Excel** section below

### Execute
The **Execute** button will run your scripts.

## How To Get Data From Excel
We use mustache rendering to tell the _**Macro Editor**_ that you want information from an excel spreadsheet.

### Example
![](/images/excel-snippet.PNG "Excel Snippet")

To set a variable (using JavaScript) from the above Excel sheet, we can do:

    `const x = {{a1:c1}};`

This will get evaluated one of two ways depending if **Cell Details** is checked.
* If **NOT CHECKED** this will evaluate to:
        
    `const x = [7, 14, 21];`

* If **CHECKED** this will evaluate to:
        
    ```
    const x = [
        {
            cell: 'a1',
            column: 'a',
            row: 1,
            value: 7
        },
        {
            cell: 'b1',
            column: 'b',
            row: 1,
            value: 14
        },
        {
            cell: 'c1',
            column: 'c',
            row: 1,
            value: 21
        }
    ];
    ```

From this point, you can interact with this variable as you would with any other variable in your desired programming language.

### What To Expect For Different Ranges
Excel allows the user to use ranges to refer to multiple cells.  For the purpose of the **Macro Editor**, you can retrieve information from an excel spreadsheet in one of three ways

> All examples below are shown with **Cell Details** value **NOT CHECKED**.

* Referring to a single cell --> `const x = {{a1}};` will evaluate to an single array of length 1 --> `const x = [7];`
* Referring to a column of numbers --> `const x = {{a1:a3}};` will evaluate to a single array of length equaling the length of the given input range--> `const x = [7, 6, 5]`
* Referring to a row of numbers --> `const x = {{a1:c1}};` will evaluate to a single array of length equaling the length of the given input range--> `const x = [7, 14, 21]`
* Referring to a matrix of numbers --> `const x = {{a1:c3}};` will evaluate to a multi dimensional array equal to the size of given input range where the inner arrays represent the rows of data-->

    ```
    const x = [
        [7, 14, 21],
        [6, 12, 18],
        [5, 10, 15],
    ];
    ```

### Things To Note
* While `a3:a1` is a valid excel range, information for that range will always be given back to the **Macro Editor** as `a1:a3`.
* This is consistent across all ranges including multi dimensional range.  Data will be retrieved from the smallest column/row to largest column row.
    * `c3:a1` and `a3:c1` will both evaluate to `a1:c3`
* Empty cells will evaluate to `""`

## How To Get Your Results To Interact With Excel
Getting Excel information into your script is only half the battle. To be useful, we now need to have our results interact back with Excel. To do this, we have created an *output* class that all supported languages have access to. This class contains three methods
* output.log
* output.write
* output.format

### The Log Method
*output.log* will log its data parameter to the output section in the editor.

* Syntax

    > output.log(*data*)

* Parameters

    _**data**_: data to be logged into output section of **Macro Editor**,  *data* can be of any data type including *boolean*, *number*, *string*, *object* and *array*

### The Write Method
*output.write* will write a set of data to a given range in the Excel spreadsheet.

* Syntax

    > output.write(*range*, *data*)

* Parameters

    _**range**_ : range for information to be written back to Excel 
    
    _**data**_ : data to be placed into given *range*.  For a single cell data should be a single number or string, For a multi cell range data should be applied as an array of numbers or strings

### The Format Method
*output.format* a given cell or range of cells with a specified color.

* Syntax

    > output.format(*range*, *color*)

* Parameters

    _**range**_ : range for cells to be formatted Excel 
    
    _**color**_ : color for cells to be highlighted with for the given *range*.  Colors can be supplied as valid hex colors: `'#FF0000'` or valid HTML color names: `'red'`

----------

#### JavaScript Notes
Due to scoping issues in JavaScript, we recommend sticking to variable = [function expression syntax](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/function), i.e. 
       
        const myFunc = function helloWorld() { 
            console.log('hello world!');
        }   
 to avoid unexepcted behavior by the interpreter. 

##Contact Us

For any issues, ideas or feedback please contact us at: <dev@burndev.co>

Developed by [![Burn Dev Logo](/images/burndev-logo.png =315x75 "Burn Dev Logo")](http://www.burndev.co)