<?xml version='1.0' encoding='UTF-8'?>
<Project Type="Project" LVVersion="26008000">
	<Property Name="NI.LV.All.SaveVersion" Type="Str">26.0</Property>
	<Property Name="NI.LV.All.SourceOnly" Type="Bool">true</Property>
	<Item Name="My Computer" Type="My Computer">
		<Property Name="NI.SortType" Type="Int">3</Property>
		<Property Name="server.app.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.control.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.tcp.enabled" Type="Bool">false</Property>
		<Property Name="server.tcp.port" Type="Int">0</Property>
		<Property Name="server.tcp.serviceName" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.tcp.serviceName.default" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.vi.callsEnabled" Type="Bool">true</Property>
		<Property Name="server.vi.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="specify.custom.address" Type="Bool">false</Property>
		<Item Name="Templates" Type="Folder">
			<Item Name="Template.xlsx" Type="Document" URL="../Templates/Template.xlsx"/>
			<Item Name="Template.xltx" Type="Document" URL="../Templates/Template.xltx"/>
			<Item Name="Template2.xlsx" Type="Document" URL="../Templates/Template2.xlsx"/>
			<Item Name="Template3.xlsx" Type="Document" URL="../Templates/Template3.xlsx"/>
		</Item>
		<Item Name="Test VIs" Type="Folder">
			<Item Name="Smoke Test.vi" Type="VI" URL="../Test/Smoke Test.vi"/>
			<Item Name="Error_Test.vi" Type="VI" URL="../Test/Error_Test.vi"/>
			<Item Name="Range Test.vi" Type="VI" URL="../Test/Range Test.vi"/>
			<Item Name="Image Test.vi" Type="VI" URL="../Test/Image Test.vi"/>
		</Item>
		<Item Name="XCel Report Engine.lvlib" Type="Library" URL="../XCel Report Engine.lvlib"/>
		<Item Name="Dependencies" Type="Dependencies"/>
		<Item Name="Build Specifications" Type="Build"/>
	</Item>
</Project>
