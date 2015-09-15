self_path =  getFilenamePath (getSourceFileName())
self_path[self_path.count] = ""
self_path+= "\\"

fn export_selection = (
	
	--filename = (maxfilepath+$.name+".obj")
	filename = "d:\\test.obj"
	exportFile filename #noPrompt selectedOnly:true using:OBJEXP 
	return filename
	)

	
fn exe_command filename= (
	
	exe_path = self_path + "main\\main.exe"
	
	command = exe_path + " i " + "\""+filename+"\""
	print command
	doscommand command
	
	)
	
	
fn exe_command_import = (
	
	exe_path = self_path + "main\\main.exe"
	
	command = exe_path + " o" 
	print command
	doscommand command
	
	)


rollout unfold_export "unfold" width:325 height:142
(
	button btn1 "导出" pos:[25,37] width:147 height:96
	

	button btn2 "导入" pos:[174,39] width:147 height:96

	on btn1 pressed  do
(
	if selection.count >= 1 do
	(
		filename = export_selection()
		exe_command filename
		
		)
	
	)
	on btn2 pressed  do
(
	
	if selection.count >= 1 do
	(
		last_select = selection as array
		for a in last_select do 
		(
			a.name = a.name+"_not_uv"
			hide a
		)
		
		exe_command_import()
		
		importfile "d:\\test.obj" #noPrompt
		
		now = selection as array
		
		group $
		)
	
	)

)
	
	
createdialog unfold_export