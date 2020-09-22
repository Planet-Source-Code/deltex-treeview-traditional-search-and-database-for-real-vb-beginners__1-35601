' BRIEF INTRO:
' ------------
'
' This is a simple tutorial for real VB beginners. I assume you have some knowledge of database to fully understand
' this tutorial, since this also deals with small amount of SQL statements. But this doesnt manipulate database
' since we will not add, edit or delete any records. We will just view it. THIS IS NOT A DATA-ENTRY.
' I also assume that the beginner knows how to make or understand a function. 
' So this tutorial is actually for OLD Programmers, who are really NEW to VB. Just like me. :)
'
' Now, about treeview, I also got this treeview concepts from many PSC coders around. they are many to mention actually.
' And if they happen to read this, i would like to thank them for posting their own which is very helpful
' to my very young 1 1/2 months self-cruisin' to VB.
'
' This sample tutorial is just a part of a complete application which i will also post in a couple of days.
'
' So LET'S BEGIN
' --------------
'
' Treeview provides a better way to search and categorize your data. Much user-friendly and intuitive. VB provides
' a way to make treeview even more interesting by attaching images to it.
'
' If you want to have images for your treeview, you must first supply an imagelist to your form. To add an imagelist,
' click on PROJECT > COMPONENTS > then make sure that Microsoft Windows Common Controls 6.0 (SP4) is checked.
'
' Then from the toolbox, click and drag an imagelist (the control with picture that looks like three white envelopes) to your form.
' Now added to the form, right-click the imagelist list, click on PROPERTIES, go to IMAGES tab and insert images that you want to use for your treeviews.
' Note that the order of images corresponds to its index. For example, in this project i have imagelist (named imgTree),
' which has 5 images in it: 
	
   Index    Image
   -----	------
	1 - Closed Folder 
	2 - Oped Folder
	3 - Text Picture
	4 - Closed catalog box
	5 - Open Catalog.
'
' Make sure that the order of the images is relevant and easy to remember, so you can you use it with ease in treeview.
' Don't worry if you dont understand it at this point, you will when you use it in code.
'
' Now, to reference the imagelist to treeview, follow this part: Select you treeview, right-click on it, click on
' PROPERTIES. 
' The Property Pages for the treeview appears. 
' Click on the GENERAL Tab. Make sure that the imagelist
' list combo box points to your own imagelist (which in this project is imgTree). 
' After you verify that this is done, you can now use the images in your treeview.
'
' To fully appreciate it, Run this project. Play around with the treeview. 
' Stop the project and go to codes page of VB. Examine the COLLAPSE
' and EXPAND method of the treeview. We use the IMAGE property of the NODE to manipulate or to reference the images
' we created in our imagelist. As ive said, the index of images in our imagelist should be relevant and easy to remember.
' In our first treeview, (By Letter treeview), when a node is COLLAPSEd, we use image 1, and when it is EXPANDed, we use image 2.
' In the second Treeview, (By Type treeview), when a node is COLLAPSEd, we use image 4, and when EXPANDed, we use image 5.
' For the data itself, we use a common index which is index 3.
'
' See the table below for better understanding:
'
'  +--------------------+--------------------+---------------------+
'  |  Treeview name     | Collapsed Image    |  Expanded Image     |
'  +--------------------+--------------------+---------------------+
'  | tvwLetters         |    1 (Open folder) |   2 (Closed Folder) |
'  | tvwType            |    4 (Open catalog)|   5 (Closed Catalog)|
'  | Data/Records       |    3 (Text picture)|   3 (Text picture)  |
'  +--------------------+--------------------+---------------------+
'
'  Run the program and view this separately to have a visual understanding of what i talked about. 
' This readme is just a portion of tutorial. you can find anything that's not written here in VB code page.
' 
' Again, thanks for downloading this.