<?xml version='1.0' encoding='UTF-8'?>
<Library LVVersion="26008000">
	<Property Name="NI.Lib.Description" Type="Str">Provides the supported LabVIEW API for creating and modifying Microsoft Excel reports through the XCelReportEngine .NET backend.
The library manages report sessions and exposes operations for opening, saving, and closing workbooks; selecting worksheets; reading and writing cells and ranges; inserting and formatting images; applying cell formatting; and protecting report contents.
Public VIs and type definitions constitute the supported client API. Private members and items contained in the DotNet section implement validation, error conversion, and communication with the backend and must not be called directly by client applications.
The library does not evaluate Excel formulas, execute VBA code, convert reports to PDF, or create charts programmatically.
Use the public VIs and type definitions as the supported API. Members located in the Private and DotNet sections are implementation details and must not be called directly by client applications.</Property>
	<Property Name="NI.Lib.HelpPath" Type="Str"></Property>
	<Property Name="NI.Lib.Icon" Type="Bin">*A#!!!!!!!)!"1!&amp;!!!-!%!!!@````]!!!!"!!%!!!(]!!!*Q(C=\&gt;8"=&gt;MQ%!8143;(8.6"2CVM#WJ",7Q,SN&amp;(N&lt;!NK!7VM#WI"&lt;8A0$%94UZ2$P%E"Y.?G@I%A7=11U&gt;M\7P%FXB^VL\`NHV=@X&lt;^39O0^N(_&lt;8NZOEH@@=^_CM?,3)VK63LD-&gt;8LS%=_]J'0@/1N&lt;XH,7^\SFJ?]Z#5P?=F,HP+5JTTF+5`Z&gt;MB$(P+1)YX*RU2DU$(![)Q3YW.YBG&gt;YBM@8'*\B':\B'2Z&gt;9HC':XC':XD=&amp;M-T0--T0-.DK%USWS(H'2\$2`-U4`-U4`/9-JKH!&gt;JE&lt;?!W#%;UC_WE?:KH?:R']T20]T20]\A=T&gt;-]T&gt;-]T?/7&lt;66[UTQ//9^BIHC+JXC+JXA-(=640-640-6DOCC?YCG)-G%:(#(+4;6$_6)]R?.8&amp;%`R&amp;%`R&amp;)^,WR/K&lt;75?GM=BZUG?Z%G?Z%E?1U4S*%`S*%`S'$;3*XG3*XG3RV320-G40!G3*D6^J-(3D;F4#J,(T\:&lt;=HN+P5FS/S,7ZIWV+7.NNFC&lt;+.&lt;GC0819TX-7!]JVO,(7N29CR6L%7,^=&lt;(1M4#R*IFV][.DX(X?V&amp;6&gt;V&amp;G&gt;V&amp;%&gt;V&amp;\N(L@_Z9\X_TVONVN=L^?Y8#ZR0J`D&gt;$L&amp;]8C-Q_%1_`U_&gt;LP&gt;WWPAG_0NB@$TP@4C`%`KH@[8`A@PRPA=PYZLD8Y!#/7SO!!!!!!</Property>
	<Property Name="NI.Lib.SourceVersion" Type="Int">637566976</Property>
	<Property Name="NI.Lib.Version" Type="Str">0.1.0.0</Property>
	<Property Name="NI.LV.All.SourceOnly" Type="Bool">true</Property>
	<Item Name="Controls" Type="Folder">
		<Property Name="NI.LibItem.Scope" Type="Int">1</Property>
		<Item Name="Border edges.ctl" Type="VI" URL="../Controls/Border edges.ctl"/>
		<Item Name="Border style.ctl" Type="VI" URL="../Controls/Border style.ctl"/>
		<Item Name="Horizontal alignment.ctl" Type="VI" URL="../Controls/Horizontal alignment.ctl"/>
		<Item Name="Image Alignment.ctl" Type="VI" URL="../Controls/Image Alignment.ctl"/>
		<Item Name="Measurement System.ctl" Type="VI" URL="../Controls/Measurement System.ctl"/>
		<Item Name="Picture Color Type.ctl" Type="VI" URL="../Controls/Picture Color Type.ctl"/>
		<Item Name="Report Ref.ctl" Type="VI" URL="../Controls/Report Ref.ctl"/>
		<Item Name="Vertical alignment.ctl" Type="VI" URL="../Controls/Vertical alignment.ctl"/>
	</Item>
	<Item Name="DotNet" Type="Folder">
		<Property Name="NI.LibItem.Scope" Type="Int">2</Property>
		<Item Name="DocumentFormat.OpenXml.dll" Type="Document" URL="../DotNet/DocumentFormat.OpenXml.dll"/>
		<Item Name="DocumentFormat.OpenXml.Framework.dll" Type="Document" URL="../DotNet/DocumentFormat.OpenXml.Framework.dll"/>
		<Item Name="XCelReportEngine.dll" Type="Document" URL="../DotNet/XCelReportEngine.dll"/>
	</Item>
	<Item Name="Private" Type="Folder">
		<Property Name="NI.LibItem.Scope" Type="Int">2</Property>
		<Item Name="Convert .NET Exception.vi" Type="VI" URL="../Private/Convert .NET Exception.vi"/>
		<Item Name="Validate Report Reference.vi" Type="VI" URL="../Private/Validate Report Reference.vi"/>
	</Item>
	<Item Name="Public" Type="Folder">
		<Property Name="NI.LibItem.Scope" Type="Int">1</Property>
		<Property Name="NI.SortType" Type="Int">3</Property>
		<Item Name="New Report.vi" Type="VI" URL="../Public/New Report.vi"/>
		<Item Name="Close Report.vi" Type="VI" URL="../Public/Close Report.vi"/>
		<Item Name="Save Report.vi" Type="VI" URL="../Public/Save Report.vi"/>
		<Item Name="Lock Report.vi" Type="VI" URL="../Public/Lock Report.vi"/>
		<Item Name="Unlock Report.vi" Type="VI" URL="../Public/Unlock Report.vi"/>
		<Item Name="Get Worksheet Names.vi" Type="VI" URL="../Public/Get Worksheet Names.vi"/>
		<Item Name="Select Worksheet.vi" Type="VI" URL="../Public/Select Worksheet.vi"/>
		<Item Name="Select Worksheet by Name.vi" Type="VI" URL="../Public/Select Worksheet by Name.vi"/>
		<Item Name="Select Worksheet by Index.vi" Type="VI" URL="../Public/Select Worksheet by Index.vi"/>
		<Item Name="Get Active Worksheet Name.vi" Type="VI" URL="../Public/Get Active Worksheet Name.vi"/>
		<Item Name="Read Cell String.vi" Type="VI" URL="../Public/Read Cell String.vi"/>
		<Item Name="Write Cell.vi" Type="VI" URL="../Public/Write Cell.vi"/>
		<Item Name="Write Cell String.vi" Type="VI" URL="../Public/Write Cell String.vi"/>
		<Item Name="Write Cell Double.vi" Type="VI" URL="../Public/Write Cell Double.vi"/>
		<Item Name="Write Cell Boolean.vi" Type="VI" URL="../Public/Write Cell Boolean.vi"/>
		<Item Name="Read String Range.vi" Type="VI" URL="../Public/Read String Range.vi"/>
		<Item Name="Write String Table.vi" Type="VI" URL="../Public/Write String Table.vi"/>
		<Item Name="Append Image.vi" Type="VI" URL="../Public/Append Image.vi"/>
		<Item Name="Format Image.vi" Type="VI" URL="../Public/Format Image.vi"/>
		<Item Name="Coordinate2Address.vi" Type="VI" URL="../Public/Coordinate2Address.vi"/>
		<Item Name="Address2Coordinate.vi" Type="VI" URL="../Public/Address2Coordinate.vi"/>
		<Item Name="Set Cell Alignment.vi" Type="VI" URL="../Public/Set Cell Alignment.vi"/>
		<Item Name="Set Cell Color and Border.vi" Type="VI" URL="../Public/Set Cell Color and Border.vi"/>
	</Item>
</Library>
