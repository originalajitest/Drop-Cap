# Drop-Cap

### Description
<ul>
	<li> This is a Word Add-in that helps automate adding drop cap to paragraphs below titles for stylized effect</li>
	<li> It has been designed primarily to work on the windows version of Word, but might work on other versions, no gurantees though </li>
	<li> To install and use, download the zip file and run the set-up inside it. It should show up in word immediately after.</li>
</ul>

### Options Descriptions
By default all descriptions are included in the tool tip for each element.
<ul>
	<li><b>Title Identifiers: </b> 
		<ul>
			<li><i>Font {1,2,3,4}:</i> Fonts of titles, all titles will be skipped and the first paragraph which is not a title will be formatted to have a Drop Cap. A blank input means that Font is not checked.</li>
			<li><i>Size:</i> The Cut-off for the lowest size that a title can be, above this it will be a title unless there is a font specified which will hence be checked.</li>
			<li><i>Divider:</i> Skips one paragraph after the title if checked to allow of a divider (----------) etc.</li>
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

