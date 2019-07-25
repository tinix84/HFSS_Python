import win32com.client


def open_ansoft_electronics_desktop():
    oAnsys = win32com.client.Dispatch("Ansoft.ElectronicsDesktop")
    oDesktop = oAnsys.GetAppDesktop()
    print('Successfully Opened Desktop App\n')
    return [oAnsys, oDesktop]


def new_maxwell3d_eddy_current(oDesktop, sDesignName: str):
    oProject = oDesktop.NewProject()
    oProject.InsertDesign("Maxwell 3D", sDesignName, "EddyCurrent", "")
    oDesign = oProject.SetActiveDesign(sDesignName)
    oEditor = oDesign.SetActiveEditor("3D Modeler")
    return [oProject, oDesign, oEditor]


def set_model_units_in(oEditor):
    oEditor.SetModelUnits(["NAME:Units Parameter", "Units:=", "in", "Rescale:=", True])


def set_model_units_mm(oEditor):
    oEditor.SetModelUnits(["NAME:Units Parameter", "Units:=", "mm", "Rescale:=", True])


if __name__ == "__main__":
    [oAnsys, oDesktop] = open_ansoft_electronics_desktop()
    [oProject, oDesign, oEditor] = new_maxwell3d_eddy_current(oDesktop, sDesignName='pippo')
    set_model_units_mm(oEditor)
