from win32com.client.gencache import EnsureDispatch, EnsureModule
from win32com.client import CastTo, constants
import matplotlib.pyplot as plt
import time
# Notes
#
# The python project and script was tested with the following tools:
#       Python 3.4.3 for Windows (32-bit) (https://www.python.org/downloads/) - Python interpreter
#       Python for Windows Extensions (32-bit, Python 3.4) (http://sourceforge.net/projects/pywin32/) - for COM support
#       Microsoft Visual Studio Express 2013 for Windows Desktop (https://www.visualstudio.com/en-us/products/visual-studio-express-vs.aspx) - easy-to-use IDE
#       Python Tools for Visual Studio (https://pytools.codeplex.com/) - integration into Visual Studio
#
# Note that Visual Studio and Python Tools make development easier, however this python script should should run without either installed.

class MtfMFGenerator(object):
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
            raise MtfMFGenerator.ConnectionException("Unable to intialize COM connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise MtfMFGenerator.InitializationException("Unable to acquire ZOSAPI application")

        if self.TheApplication.IsValidLicenseForAPI == False:
            raise MtfMFGenerator.LicenseException("License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise MtfMFGenerator.SystemNotPresentException("Unable to acquire Primary system")

    def __del__(self):
        """Boiler plate"""
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None

        self.TheConnection = None

    def OpenFile(self, filepath, saveIfNeeded):
        """Boiler plate"""
        if self.TheSystem is None:
            raise MtfMFGenerator.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)

    def CloseFile(self, save):
        """Boiler plate"""
        if self.TheSystem is None:
            raise MtfMFGenerator.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.Close(save)

    def SamplesDir(self):
        """Boiler plate"""
        if self.TheApplication is None:
            raise MtfMFGenerator.InitializationException("Unable to acquire ZOSAPI application")

        return self.TheApplication.SamplesDir

    def ExampleConstants(self):
        """Boiler plate"""
        if self.TheApplication.LicenseStatus is constants.LicenseStatusType_PremiumEdition:
            return "Premium"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_ProfessionalEdition:
            return "Professional"
        elif self.TheApplication.LicenseStatus is constants.LicenseStatusType_StandardEdition:
            return "Standard"
        else:
            return "Invalid"

    def RemoveAllAfterDMFS(self):
        """Remove all the oparands after the first DMFS in the merit function editor"""
        mfe = self.TheSystem.MFE
        nRows = mfe.NumberOfOperands
        dmfsrow = -1
        for r in range(nRows):
            row = mfe.GetOperandAt(r)
            if(row.Type == constants.MeritOperandType_DMFS):
                dmfsrow = r
                break
        if(dmfsrow > 0):
            CastTo(mfe,'IEditor').DeleteRowsAt(dmfsrow, nRows - dmfsrow)

    def AddMTFOPGT(self, field, freq, target,type):
        """
        Add operand of type type (GMTS or GMTT) for field and freq. 
        Then add operant OPGT requiring the previous operand to be larger than target
        """
        mfe = self.TheSystem.MFE
        mtf = mfe.AddOperand()
        mtf.ChangeType(type)
        p1 = mtf.GetOperandCell(constants.MeritColumn_Param1)
        p1.IntegerValue = 2
        p3 = mtf.GetOperandCell(constants.MeritColumn_Param3)
        p3.IntegerValue = field + 1
        p4 = mtf.GetOperandCell(constants.MeritColumn_Param4)
        p4.DoubleValue = freq
        p6 = mtf.GetOperandCell(constants.MeritColumn_Param6)
        p6.IntegerValue = 1
        
        opgt = mfe.AddOperand()
        opgt.ChangeType(constants.MeritOperandType_OPGT)
        opgt.Target = target
        wc = opgt.GetOperandCell(constants.MeritColumn_Weight)
        wc.DoubleValue = 1.0
        p1 = opgt.GetOperandCell(constants.MeritColumn_Param1)
        p1.IntegerValue = CastTo(mtf,'IEditorRow').RowIndex + 1
        
    def OptimizeMTFGreaterThan(self, nFields, freq, target):
        """
        Create MF to optimize on MTF for the nFields first field points, trying to make it greater than target at freq. 
        Operands will be added to the end of the merit function.
        """
        mce = self.TheSystem.MCE
        mcs = mce.NumberOfConfigurations
        for mc in range(mcs):
            mfe = self.TheSystem.MFE
            cnf = mfe.AddOperand()
            cnf.ChangeType(constants.MeritOperandType_CONF)
            cnf.GetOperandCell(constants.MeritColumn_Param1).IntegerValue = mc + 1
            for f in range(nFields):
                self.AddMTFOPGT(f, freq, target, constants.MeritOperandType_GMTS)
                self.AddMTFOPGT(f, freq, target, constants.MeritOperandType_GMTT)

    def LocalOptimizeMTF(self, target):
        """
        Start local optimization, keep tunning until MF is below target, local optimization converges, 
        or 500 minutes has passed.
        """
        lopt = self.TheSystem.Tools.OpenLocalOptimization()
        lopt.Algorithm = constants.OptimizationAlgorithm_DampedLeastSquares
        lopt.Cycles = constants.OptimizationCycles_Infinite
        lopt.NumberOfCores = 8
        print("Starting lopt")    
        CastTo(lopt, "ISystemTool").Run()
        mf = lopt.InitialMeritFunction
        counter = 0
        dcount = 0
        print("Starting loop, mf = " + str(mf))
        while mf > target:
            time.sleep(60)
            if (lopt.CurrentMeritFunction < mf):
                dcount = 0

            if dcount > 0: print("dcount is " + str(dcount))    
            mf = lopt.CurrentMeritFunction
            print("mf = " + str(mf))
            counter = counter + 1
            dcount = dcount + 1
            if( counter > 500): break
            if( dcount > 5): break
        CastTo(lopt, "ISystemTool").Cancel()
        CastTo(lopt, "ISystemTool").Close()
        return(mf)

    def HammerOptimize(self, target):
        """
        Start hammer optimization. Keep running until MF is below target, printing current MF every 10 minutes.
        """

        hopt = self.TheSystem.Tools.OpenHammerOptimization()
        hopt.Algorithm = constants.OptimizationAlgorithm_DampedLeastSquares
        hopt.NumberOfCores = 8
        print("Starting hopt")    
        CastTo(hopt, "ISystemTool").Run()
        mf = hopt.InitialMeritFunction
        print("Starting loop, mf = " + str(mf))
        iter = 0
        while mf > target:
            time.sleep(600)
            mf = hopt.CurrentMeritFunction
            print("Time " + str(iter))
            print("mf = " + str(mf))
            iter = iter + 1
        CastTo(hopt, "ISystemTool").Cancel()
        CastTo(hopt, "ISystemTool").Close()
        return(mf)
    
def OptimizeMTF(target, maxfreq, startfreq):
    """Optimize on MTF for increasing frequency, first using local optimization, then hammer.

    Merit function requires GMTS and GMTT to be above 0.5 for all
    points of view. When this is achieved, we try again for a
    frequency 0.25 higher, until maxfreq.

    Starting Hammer optimization after local optimization hangs the ZOSAPI connection, I do not know why. 
    for This reason, the ZOSAPI connection is opened and closed for every optimizer.

    Kill with ctrl+c in powershell
    """
    freq = startfreq
    while freq <= maxfreq:
        #Set up for local optimization
        print('Preparing for freq ' + str(freq))
        zosapi = MtfMFGenerator()
        value = zosapi.ExampleConstants()
        zosapi.OpenFile('m:\\tmp2.zmx',False)
        zosapi.RemoveAllAfterDMFS()
        zosapi.OptimizeMTFGreaterThan(5, freq, 0.5)
        mf = zosapi.LocalOptimizeMTF(target)
        zosapi.TheSystem.SaveAs("m:\\tmp2.zmx")
        del zosapi
        print('MF after local optimization is ' + str(mf))

        #Global optimization
        zosapi = MtfMFGenerator()
        value = zosapi.ExampleConstants()
        zosapi.OpenFile('m:\\tmp2.zmx',False)
        zosapi.RemoveAllAfterDMFS();
        zosapi.OptimizeMTFGreaterThan(5, freq, 0.5)
        zosapi.HammerOptimize(target)
        zosapi.TheSystem.SaveAs("m:\\tmp2.zmx")
        del zosapi
        freq = freq + 0.25
        
if __name__ == '__main__':
    #Make sure paths are ok before running
    # Insert Code Here
    # Open file
    OptimizeMTF(0.001, 13, 4.25)

