import win32com.client


class Maxwell_3D():
    def __init__(self):
        self.oAnsys = None
        self.oDesktop = None
        self.oProject = None
        self.oDesign = None
        self.oEditor = None
        self.oModule = None
        self.design_name = None

    def open_ansoft_electronics_desktop(self):
        self.oAnsys = win32com.client.Dispatch("Ansoft.ElectronicsDesktop")
        self.oDesktop = self.oAnsys.GetAppDesktop()
        print('Successfully Opened Desktop App\n')

    def new_maxwell3d_eddy_current(self, sDesignName: str):
        self.design_name = sDesignName
        self.oProject = self.oDesktop.NewProject()
        self.oProject.InsertDesign("Maxwell 3D", self.design_name, "EddyCurrent", "")
        self.oDesign = oProject.SetActiveDesign(self.design_name)
        self.oEditor = oDesign.SetActiveEditor("3D Modeler")

    def set_model_units_in(self):
        self.oEditor.SetModelUnits(["NAME:Units Parameter",
                                    "Units:=", "in",
                                    "Rescale:=", True])

    def set_model_units_mm(self):
        self.oEditor.SetModelUnits(["NAME:Units Parameter",
                                    "Units:=", "mm",
                                    "Rescale:=", True])

