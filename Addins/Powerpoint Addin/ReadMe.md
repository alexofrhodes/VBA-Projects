
To Show Addin Projects in the VBE Window you have to Modify Your Registry

* Close down PowerPoint

* Go to your Start Menu, Type in Regedit.exe and click OK

* Navigate to the following key in the registry tree: 
	HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options		
	(16.0 may be a different number based on your Office version)

* Find or create the key name DebugAddins and Set the DWORD value to 1
	 ________/_________________________   
	 Edit > New > DWORD (32b-it)) Value

* Launch PowerPoint and go into your Visual Basic Editor (Alt+F11).  
  You will now be able to see any PowerPoint add-in VBA code that is currently running


If you modify the addin project's vbproject (modules / procedures) the changes will NOT be saved.
You have to keep the original .PPTM, modify that after disabling the addin, overwrite the addin and reenable.
This Tool offers a userform facilitate the addin editing (@TODO add xml to created .PPTM) 
or simply export the components.
 
[![PowerPointAddin](https://img.youtube.com/vi/oPLJNNdK_bc/0.jpg)](https://www.youtube.com/watch?v=oPLJNNdK_bc)


