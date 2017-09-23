//    This file is part of DotNetInteropDemos
//    Copyright (C) James Forshaw 2017
//
//    DotNetInteropDemos is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    DotNetInteropDemos is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with DotNetInteropDemos.  If not, see <http://www.gnu.org/licenses/>.

try {
	var ax_vb = new ActiveXObject("Microsoft.Windows.ActCtx");
	ax_vb.Manifest = "vb.manifest";
	var ax_al = new ActiveXObject("Microsoft.Windows.ActCtx");
	ax_al.Manifest = "arraylist.manifest";

	var obj = ax_vb.CreateObject("Microsoft.VisualBasic.Devices.Computer");
	var clipboard = obj.Clipboard;
	
	var appbase = ax_vb.CreateObject("Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase")
	var asms = ax_al.CreateObject("System.Collections.ArrayList");
	asms.AddRange(appbase.Info.LoadedAssemblies);
	var mscorlib = asms.Item(0);
	var args = ax_al.CreateObject("System.Collections.ArrayList");
	var empty_array = ax_al.CreateObject("System.Collections.ArrayList");
	var null_obj = appbase.SplashScreen;
	args.Add(false);

	var ev = mscorlib.CreateInstance_3("System.Threading.AutoResetEvent", false, 4 | 16, 
			null_obj, args.ToArray(), null_obj, empty_array.ToArray());

	var last_clipboard = clipboard.GetText();
	while(true) {
		ev.WaitOne_4(1000);
		var text = clipboard.GetText();
		if (text != last_clipboard) {
			WScript.Echo("Captured: " + text);
			last_clipboard = text;
		}
	}
} catch(e) {
	WScript.Echo("Error: " + e.message);
}