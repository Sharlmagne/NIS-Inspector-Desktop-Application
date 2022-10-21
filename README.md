# NIS Inspector Application

NIS Inspector is a desktop application created to simplify the tedious tasks of manual data entry that clerks have to perform on a weekly basis. Where did it begin? Well, my mom is an inspector for the National Insurance Scheme in Jamaica. This is a compulsory contributory funded social security scheme, by the government of Jamaica, that covers all employed persons in Jamaica. Her role involves going around and ensuring that employers make contributions to the scheme among other things. On a weekly basis, she is tasked with forty different objectives that she will have to manually write on a form (itinerary). These objectives include the date, name of the business, address of the business, type of visit, etc. By the end of the week, she will have to manually re-enter all the information on another form (report) with additional details to describe the outcome of her objectives. Manually writing 80 different objectives on tons of paper screams inefficiency, a waste of resources, and a missed opportunity to collect valuable data. So, this is where I came in to save the data for my mom and her office, and hopefully the entire organization and possibly the entire country... Maybe I'm getting ahead of myself. Nonetheless, NIS Inspector was born to simplify the process and usher the government into the digital age. NIS Inspector was created using Python and the PyQt module to develop a GUI to quickly import information, store it as json files and automatically transfer the data to word documents in no time. Fueled by love and a knack for problem-solving, I'm now my mother's favorite child by making her life easier.

## Files
- executable - contains the setup executable and the script to create the setup file using the Inno Compiler
- dist - this is the folder that the executable will load onto your computer
- ui - all the ui elements
- files - source documents used in the program along with the auto save states
- venv - virtual environment for python 3.10.6... I believe :)

## Source code
- nisi.py - main file
- data.py - data prep for the program
- logo_rc.py - load the logo to the gui
- report.py - class for the report
- itinerary.py class for the itinerary
- ui.py class for the gui including the bulk of the methods used to operate the program

## Snapshots
![itinerary_tab](/ui/snapshot-1.png)
![report_tab](/ui/snapshot-2.png)
