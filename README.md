# LoL-Spreadsheet
An Excel spreadsheet for logging League of Legends matches. Main feature is a UserForm which eases logging matches.

Features
---------
* Form with fields for each column which makes entering into the spreadsheet directly unnecessary
	* Automatically calculates and enters:
    	* Date
    	* Net LP change
    		* Can account for dodged games/negative LP
    	* CS per minute
    	* Gold per minute
* Many options for:
   	* Which fields to clear
   	* Which fields can be navigated with the tab key
* Conditionally formatted columns

#### Planned Features
* C# port
* Improve conditionally formatted columns
    * Different ranges per role <sup id="r1">[1](#fn1)</sup>
    * Differents ranges per champion <sup id="r2">[1](#fn2)</sup>
* Add error handling!!!
* Importing match data stored on a file

Instructions
---------

### Requirements

So far, the VBA macro has only been tested on:

* Excel 2013
* Excel from Office 365 ProPlus 2016.

I will update this list if I get confirmation of additional working software.

### Usage

Upon opening the spreadsheet for the first time, you may be given prompts on yellow horizontal bars at the top of the screen. One will warn that the file is from the internet and prompt you to enable editing. Do so. The other will prompt you to enable macros, do so as well.

To log a match, click on the button named __Enter New Match__. A UserForm will appear. Here is where you enter all of the details for the match. Note that some fields, such as the *Screenshot* and the *Dodged* checkboxes, are optional.

Once you have completed the form, you can transfer the data to the spreadsheet by pressing the __Submit__ button. If a mandatory field was entered incorrectly or left empty, the macro will mark that field in red for you (this can be partially disabled.)

### Configuration
The macro has several quite specific settings for specifying which fields to clear and which fields to skip over when navigating with the tab key. These settings are located in the *Misc. Settings* tab along with those for disabling error checking. The error checking settings cannot disable error checking for fields which must be numerical.<sup id="r3">[3](#fn3)</sup>


Known Issues
---------
* Saving settings to a file is disabled
    * Original implementation idea didn't work; I do have some ideas as to how to implement it
* Scroll wheel does not function on ComboBoxes (drop down lists)
	* Seems unlikely I can make this word but I will keep digging around

Credits
------
* Mark Kozlov
* [Stack Overflow](http://stackoverflow.com)
    * [PsychoCoder](http://stackoverflow.com/users/88230/psychocoder) - http://stackoverflow.com/a/3426721/5717792
    * [Matteo](http://stackoverflow.com/a/32119820/5717792) - http://stackoverflow.com/a/32119820/5717792
    * [statue](http://stackoverflow.com/users/833921/statue) - http://stackoverflow.com/a/20661481/5717792
---

###### Footnotes
<a name="fn1">1</a> Because, for example, certain stats like gold and CS per minute have much lower averages for supports and junglers. [↩](#r1)
<a name="fn2">2</a> Mutch more abitious, but it's something I think is possible and want to try. [↩](#r2)
<a name="fn3">3</a> This is to prevent crashes because I'm too lazy to check if everything is numerical, so I just force only numerical inputs. [↩](#r3)

