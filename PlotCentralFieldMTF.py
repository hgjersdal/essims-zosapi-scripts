from win32com.client.gencache import EnsureDispatch, EnsureModule
from win32com.client import CastTo, constants
import matplotlib.pyplot as plt
import time
from array import array
from math import floor
# Notes
#
# The python project and script was tested with the following tools:
#       Python 3.4.3 for Windows (32-bit) (https://www.python.org/downloads/) - Python interpreter
#       Python for Windows Extensions (32-bit, Python 3.4) (http://sourceforge.net/projects/pywin32/) - for COM support
#       Microsoft Visual Studio Express 2013 for Windows Desktop (https://www.visualstudio.com/en-us/products/visual-studio-express-vs.aspx) - easy-to-use IDE
#       Python Tools for Visual Studio (https://pytools.codeplex.com/) - integration into Visual Studio
#
# Note that Visual Studio and Python Tools make development easier, however this python script should should run without either installed.
class Histos:
    def __init__(self):
        self.resolutions = []
        self.histos5 = [[],[],[]]
        self.histos75 = [[],[],[]]
        self.histos10 = [[],[],[]]

    def FillCounterHisto(self, histos, counter):
        for i in range(3):
            histos[i].append(counter[i])

    def PlotHistos(self, path, bname, histos):
        typeNames = ["tangential","sagittal","average"]
        for i in range(3):
            fig,ax = plt.subplots(1,1,figsize=(8,6))
            plt.hist(histos[i])
            plt.grid()
            fig.savefig(path + bname + '-' + typeNames[i] + '.png')
            plt.close(fig)
            
class PlotCentralFieldMTF(object):

    class LicenseException(Exception):
        pass

    class ConnectionException(Exception):
        pass

    class InitializationException(Exception):
        pass

    class SystemNotPresentException(Exception):
        pass

    def __init__(self):
        # make sure the Python wrappers are available for the COM client and
        # interfaces
        EnsureModule('ZOSAPI_Interfaces', 0, 1, 0)
        # Note - the above can also be accomplished using 'makepy.py' in the
        # following directory:
        #      {PythonEnv}\Lib\site-packages\wind32com\client\
        # Also note that the generate wrappers do not get refreshed when the
        # COM library changes.
        # To refresh the wrappers, you can manually delete everything in the
        # cache directory:
        #	   {PythonEnv}\Lib\site-packages\win32com\gen_py\*.*
        
        self.TheConnection = EnsureDispatch("ZOSAPI.ZOSAPI_Connection")
        if self.TheConnection is None:
            raise PlotCentralFieldMTF.ConnectionException("Unable to intialize COM connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise PlotCentralFieldMTF.InitializationException("Unable to acquire ZOSAPI application")

        if self.TheApplication.IsValidLicenseForAPI == False:
            raise PlotCentralFieldMTF.LicenseException("License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise PlotCentralFieldMTF.SystemNotPresentException("Unable to acquire Primary system")

    def __del__(self):
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None

        self.TheConnection = None

    def OpenFile(self, filepath, saveIfNeeded):
        if self.TheSystem is None:
            raise PlotCentralFieldMTF.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)

    def CloseFile(self, save):
        if self.TheSystem is None:
            raise PlotCentralFieldMTF.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.Close(save)

    def SamplesDir(self):
        if self.TheApplication is None:
            raise PlotCentralFieldMTF.InitializationException("Unable to acquire ZOSAPI application")

        return self.TheApplication.SamplesDir

    def ExampleConstants(self):
        if self.TheApplication.LicenseStatus is constants.LicenseStatusType_PremiumEdition:
            return "Premium"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_ProfessionalEdition:
            return "Professional"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_StandardEdition:
            return "Standard"
        else:
            return "Invalid"

    def RemoveExtremeFields(self):
        """Remove field points 6,7,8,9 and 1. """
        field = self.TheSystem.SystemData.Fields
        for x in range (9,5,-1):
            field.RemoveField(x)
        field.RemoveField(1)

    def CheckLimits(self, xdata, ydata, index, histos):
        resolution = 20.0
        for i in range(xdata.Length):
            if ydata.Data[i][index] < 0.25:
                lim = xdata.Data[i]
                resolution = xdata.Data[i]
                break
        histos.resolutions.append(20)
        return(resolution)

    def CornerCounter (self,res, index, h5, h75, h10):
        if res > 5.0:
            h5[index] = 1 + h5[index]
        if res > 7.5:
            h75[index] = 1 + h75[index]
        if res > 10.0:
            h10[index] = 1 + h10[index]
            
    def PlotMtfAllConfigs(self, bname, histos):
        """Loop over all configs in MCE, and plot the MTF for all active fields"""
        mce = self.TheSystem.MCE
        mcs = mce.NumberOfConfigurations
        #Loop over all configs
        for mc in range(mcs):
            mce.SetCurrentConfiguration(mc + 1)
            
            #Plot MTF
            gmtf = self.TheSystem.Analyses.New_GeometricMtf()
            settings = CastTo( gmtf.GetSettings(), 'IAS_GeometricMtf' )
            #settings.ShowDiffractionLimit()
            settings.MaximumFrequency = 20.0
            gmtf.ApplyAndWaitForCompletion()
            #gmtf.ToFile('m:\\gmtf.txt')
            results = gmtf.GetResults()
            
            fig, ax = plt.subplots(1,1, figsize=(8,6))
            res5  = [0,0,0]
            res75 = [0,0,0]
            res10 = [0,0,0]

            #Loop over results.
            for i in range(results.NumberOfDataSeries):
                ds = results.GetDataSeries(i)
                plt.plot(ds.XData.Data,ds.YData.Data)
                resT = self.CheckLimits(ds.XData, ds.YData, 0, histos) #Tangential (or opposite)
                resS = self.CheckLimits(ds.XData, ds.YData, 1, histos) #Sagittal(or opposite)
                self.CornerCounter (resT, 0, res5, res75, res10)
                self.CornerCounter (resS, 1, res5, res75, res10)
                self.CornerCounter ((resT + resS)/2.0, 2, res5, res75, res10)
                
            histos.FillCounterHisto(histos.histos5 , res5)
            histos.FillCounterHisto(histos.histos75, res75)
            histos.FillCounterHisto(histos.histos10, res10)
            plt.grid()
            fig.savefig('c:\\Users\\haavagj\\plots\\' + bname  + str(mc) + '.png')
            plt.close(fig)    
        
if __name__ == '__main__':
    """Reads file m:/tmp2.zmx, removes fields and plots the MTF for the central fields
    Make sure the paths for the plots and the input file are ok before running"""
    histos = Histos()
    print(histos.resolutions)

    for i in range(0,100):
        print('MC-alignment' + str(i))
        zosapi = PlotCentralFieldMTF()
        value = zosapi.ExampleConstants()
        zosapi.OpenFile('c:\\Users\\haavagj\\MC-alignment' + str(i) + '.zmx',False)
        zosapi.RemoveExtremeFields()
        zosapi.PlotMtfAllConfigs('mtf' + str(i), histos)
    
        # This will clean up the connection to OpticStudio.
        # Note that it closes down the server instance of OpticStudio, so you for maximum performance do not do
        # this until you need to.
        del zosapi
    histos.PlotHistos('c:\\Users\\haavagj\\plots\\', "corners-mtf5",  histos.histos5)
    histos.PlotHistos('c:\\Users\\haavagj\\plots\\', "corners-mtf75", histos.histos75)
    histos.PlotHistos('c:\\Users\\haavagj\\plots\\', "corners-mtf10", histos.histos10)
    fig,ax = plt.subplots(1,1,figsize=(8,6)) 
    plt.hist(histos.resolutions)
    plt.grid()
    fig.savefig('c:\\Users\\haavagj\\plots\\mtf-resolution.png')
    plt.close(fig)
    
