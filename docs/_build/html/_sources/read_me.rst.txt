Read Me
=======

**Repository description:**
---------------------------
	If you clone this repository to a local machine you will be able to run the Upload Assistant from uploadassistant\\main.py or run the application in Windows by starting the file uploadassistant\\dist\\Upload Assistant.exe

	If you need to create the .exe version that was sent November 2020 as part of the data collection you must set up a conda environment as explained at the bottom of this document.

**Folder structure with descriptions:**
---------------------------------------

+----------------------------------------------------------+-------------------------------------------------+
|                                                          |                                                 |
|                                                          |  System folder, do not edit here.               |
|    **/.idea**                                            |                                                 |
+----------------------------------------------------------+-------------------------------------------------+
|                                                          |                                                 |
|                                                          | The project documentation files built in sphinx,|
|                                                          | launch with file docs\\build\\html\\index.html. |
|     **/docs**                                            |                                                 |
+----------------------------------------------------------+-------------------------------------------------+
|                                                          | Contains a few files removed from the final     |
|     **/ParkingLot**                                      | project. Importing from different documents     |
|                                                          | added complexity to the creation of the .exe    |
|                                                          | version, so some of these files were            |
|                                                          | incorperated in main.py                         |
+----------------------------------------------------------+-------------------------------------------------+
|                                                          |                                                 |
|                                                          | The main project folder. you can run the final  |
|                                                          | version of the file from \\uploadassistant\\    |
|    **/uploadassistant**                                  | dist\\Upload Assistant.exe of view or run a     |
|                                                          | script version from uploadassistant\\main.py.   |
|                                                          |                                                 |
+----------------------------------------------------------+-------------------------------------------------+

**/uploadassistant content descriptions:**
------------------------------------------

+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|                                                          | System folder, do not edit.                                                                       |
|    **/.ipynb_checkpoints**                               |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|                                                          | Discarded files.                                                                                  |
|                                                          |                                                                                                   |
|     **/Archive**                                         |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|     **/dist**                                            | Contains pyinstaller output (working .exe).                                                       |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/Images**                                           | Contains images used in application and script.                                                   |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/Output**                                           | Location where the script and application will create json output files for GPM to load into      |
|                                                          | the database.                                                                                     |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/QA Resources**                                     | Files a developer can use for testing in a running script/application.                            |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/Report**                                           | Location where the script/application will create a user report of summary data.                  |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/Resources**                                        | Resources to aid the Upload Assistant end user in creation of the data collection submission, this|
|                                                          | contains folders that are custom to certain countries.                                            |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/Submission**                                       | Location where the user saves completed version of thier template to be evaluated by the          |
|                                                          | Upload Assistant.                                                                                 |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/Template**                                         | The starting version of the document to be completed and saved in the submission folder.          |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+
|                                                          |                                                                                                   |
|    **/__pycache__**                                      | System folder, do not edit.                                                                       |
|                                                          |                                                                                                   |
+----------------------------------------------------------+---------------------------------------------------------------------------------------------------+

conda env instructions:
-------------------------

In order to create a stable exe application with pyinstaller it was found to be necessary to create a very specific environment with a combination of package versions. If you do not recreate the environment exactly, make sure you are testing that your exe file workd on other users' laptops. A known issue with the venv and up-to-date version as of October 2020 is a working application that will not open on other users' computers.

These instructions are for PyCharm Community. If you edit python in another IDE you may need to adjust these.

**1.**
	Follow these instructions to setup a conda based Python 3.6 environment in any folder you choose. https://uoa-eresearch.github.io/eresearch-cookbook/recipe/2014/11/20/conda/  - if link is broken, find instructions by googling "set up a python 3.6 conda environment" after cloning this GitLab repository https://gitlab.gda.allianz/azpatravel/GHARMO.gitlab.io.git , open the folder "uploadassistantproj" as a new project in PyCharm.

	Do not create a new virtual environment. 

**2.**
	Open \uploadassistantproj\\uploadassistant\\main.py in the editor.
	Edit the configuration (button at top right).
	*For detailed PyCharm instructions reference https://www.jetbrains.com/help/pycharm/quick-start-guide.html. 
**3.**
	Add a new configuration ("+" sign on menu), select the conda environment you created.
	Make sure the working directory is [replace with your local address]\uploadassistantproj\uploadassistant
	Click ok and apply and make sure that the configuration can read files. (run a sample file that says print("hello") from editor, etc)
**4.**
	From the terminal install these packages one at a time. (Do not use the standard venv project interpreter menu)
	::

			-pip install Cerberus
			-pip install pywin32
			-pip install mouse
			-pip install numpy
			-pip install pandas
			-pip install pandastable
			-python -m pip install matplotlib==3.0.3
			-pip install pyinstaller==3.6
			-pip install nicexcel

	If there are any missing packages when you run main.py install them in the same method.
	If main.py launches, your environment is working properly.

pyinstaller creation of exe
---------------------------

**1.**
	From the console run this line: 
	::
	
	[replace with your local address]uploadassistantproj\uploadassistant>  pyinstaller --onefile -w -F -i "UAclipboard.ico"  "[replace with your local address]uploadassistantproj\uploadassistant\main.py"
	
	this will create a new file in uploadassistantproj\uploadassistant\dist\ (or it may error out, that's ok!)
	The file that is created, it may have features that are not wanted, or it may be a directory with many files, or the process may have thrown a recursion error.
**Here is what to do next.**
**2.**
	find the file "uploadassistantproj\uploadassistant\main.spec" (Not main.py!!)
	open this link and compare the two files
	https://gitlab.gda.allianz/azpatravel/GHARMO.gitlab.io/-/blob/master/uploadassistant/main.spec
	Take the custom options from the GitLab version and apply them to your default options. Start with the first 2 lines, they well prevent 
			the recursion limit error.
**3.**
	Finally, in the console run the line [Only change from the last time is it ends with ".spec" vs ".py"]:
	::
	
	[replace with your local address]uploadassistantproj\uploadassistant>  pyinstaller --onefile -w -F -i "UAclipboard.ico"  "[replace with your local address]uploadassistantproj\uploadassistant\main.spec"
			
If everything has been done correctly you will have created an application in the \dist folder that will run on a standard Windows computer.



