# Drop-Cap

## Version: &emsp;&emsp; 1.0.0.0

## Description
<ul>
	<li> This is a Word Add-in that helps automate adding drop cap to paragraphs below titles for stylized effect</li>
	<li> It has been designed primarily to work on the windows version of Word, but might work on other versions like on mac and linux, but no gurantees though </li>
	<li> To install and use, see below.</li>
	<li>It takes approximately 1 minute as or version 1.0.0.0 to drop everything and the same amount of time to clear all Dropped Characters everything</li>
</ul>

## Installing, Updating & Uninstalling

<ul>
	<li> <b>Installing</b>
		<ol>
			<li>Download the zip file (<i>Drop-Cap-Add-In.zip</i>) from the master branch root</li>
			<li>Extract it and run the setup.exe file</li>
			<li>It should now be in word whenever you open word</li>
			<li>Do not delete the folder which contains these files, it runs from there</li>
		</ol>
	</li>
	<li> <b>Updating</b>
		<ol>
			<li>You will be able to see the version number from the top of this ReadMe document</li>
			<li>Download the same zip file from root</li>
			<li>Extract it and locate the version that you want inside the folder <i>Application Files</i></li>
			<li>Take that version and paste it in the same folder as your current version (in <i>Application Files</i>)</li>
			<li><b>This can be skipped. </b>Paste the .vsto file in the new version in the main folder replacing the one there</li>
			<li>You can delete the old version if you want. Or you could simply uninstall and reinstall the new version</li>
		</ol>
	</li>
	<li> <b>Uninstalling</b>
		<ol>
			<li>Open Control Panel from the start menu</li>
			<li>Go to <i>Uninstall A Program</i> under <i>Programs</i></li>
			<li>Locate the Program with the name <i>Drop-Cap</i></li>
			<li>Right click it then click uninstall</li>
			<li>It will remove itself from word</li>
		</ol>
	</li>
</ul>

## Options Descriptions
By default all descriptions are included in the tool tip for each element.
<ul>
	<li><b>Title Identifiers: </b> 
		<ul>
			<li><i>Font {1,2,3,4}:</i> Fonts of titles, all titles will be skipped and the first paragraph which is not a title will be formatted to have a Drop Cap. A blank input means that Font is not checked.</li>
			<li><i>Size:</i> The Cut-off for the lowest size that a title can be, above this it will be a title unless there is a font specified which will hence be checked.</li>
			<li><i>Divider:</i> Skips one paragraph after the title if checked to allow for a divider (----------) etc.</li>
		</ul>
	</li>
	<li><b>Main: </b>
		<ul>
			<li><i>Auto Drop Cap:</i> Drops paragraphs according toe the identifiers. Each drop is from the beginning of the paragraph to the first alphabet.</li>
			<li><i>Remove Drop Cap:</i> Removes all Drop Cap in the document regardless of whether it was applied by the add-in or manually.</li>
		</ul>
	</li>
	<li><b>Drop Settings: </b> 
		<ul>
			<li><i>Font :</i> Sets the font of the Dropped text. Blank means the same font as the paragraph.</li>
			<li><i>Lines To Drop:</i> Lines the Dropped text should cover, default is 3.</li>
			<li><i>In Margin:</i> If the Dropped text should end before the margin or if they should be in text. Check = End before Margin.</li>
		</ul>
	</li>
</ul>

## Using the Add-In

<ul>
	<li>Load the file you want to Drop Cap into Word</li>
	<li>In <b><i>Size</i></b>, put the largest size such that all titles are larger than it (including all text which should not be dropped). If the paragraphs are the same size, ensure that they are in a different font</li>
	<li>For anything that you want the add-in to skip, put it in a different font</li>
	<li>Paste the main title font name in <b><i>Font 1</i></b>, or a substring of it if you want</li>
	<li>Paste any necessary font names as you require into <b><i>Font {2,3,4}</i></b>. The order does not matter</li>
	<li>If there is a divider of some sort (*'s or -'s etc.) below each title, check the divider box</li>
	<li>Set <b><i>Lines to Drop</i></b> as per how many lines you require the Dropped Character to occupy</li>
	<li>Set <b><i>Font</i></b> based on if you want a different font for the Dropped Character</li>
	<li>Check <b><i>In Margin</i></b> if you want the Dropped Character to end before the margin and leave it unchecked if it should be inline with the text</li>
	<li>Now simply click <b><i>Auto Drop Cap</i></b> and watch it do it's work</li>
</ul>


## Reasons for Errors

<ul>
	<li>The add-in dropped paragraphs that it should have skipped
		<ul>
			<li>There is probably a space at the start of a few paragraphs that were not formatted correctly, select the text again and give it the same font</li>
			<li>Word does not like to keep the font at the front of paragraphs for some reason ¯\_(ツ)_/¯</li>
		</ul>
	</li>
</ul>

Thats about what I have, hope it is useful :D