from win32com.client.gencache import EnsureDispatch, EnsureModule
from win32com.client import CastTo, constants
import matplotlib.pyplot as plt
import time
import random
# Notes
#
# The python project and script was tested with the following tools:
#       Python 3.4.3 for Windows (32-bit) (https://www.python.org/downloads/) - Python interpreter
#       Python for Windows Extensions (32-bit, Python 3.4) (http://sourceforge.net/projects/pywin32/) - for COM support
#       Microsoft Visual Studio Express 2013 for Windows Desktop (https://www.visualstudio.com/en-us/products/visual-studio-express-vs.aspx) - easy-to-use IDE
#       Python Tools for Visual Studio (https://pytools.codeplex.com/) - integration into Visual Studio
#
# Note that Visual Studio and Python Tools make development easier, however this python script should should run without either installed.

class MisAlignmentGenerator(object):
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
            raise MisAlignmentGenerator.ConnectionException("Unable to intialize COM connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise MisAlignmentGenerator.InitializationException("Unable to acquire ZOSAPI application")

        if self.TheApplication.IsValidLicenseForAPI == False:
            raise MisAlignmentGenerator.LicenseException("License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise MisAlignmentGenerator.SystemNotPresentException("Unable to acquire Primary system")

    def __del__(self):
        """Boiler plate"""
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None

        self.TheConnection = None

    def OpenFile(self, filepath, saveIfNeeded):
        """Boiler plate"""
        if self.TheSystem is None:
            raise MisAlignmentGenerator.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)

    def CloseFile(self, save):
        """Boiler plate"""
        if self.TheSystem is None:
            raise MisAlignmentGenerator.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.Close(save)

    def SamplesDir(self):
        """Boiler plate"""
        if self.TheApplication is None:
            raise MisAlignmentGenerator.InitializationException("Unable to acquire ZOSAPI application")

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

    def RemoveAllMtfRows(self):
        """Remove all the oparands in the merit function editor"""
        mfe = self.TheSystem.MFE
        nRows = mfe.NumberOfOperands
        dmfsrow = -1
        CastTo(mfe,'IEditor').DeleteRowsAt(0, nRows)

    def AddREAOp(self, surf, missCenter, missPupilX, missPupilY, REAXp):
        """
        Add operand of type REAX or REAY to MFE. This means how much the (almost)chief ray is going to miss the vertex of the surface.
surf is the surface that we are aiming for
missCenter is the amount we miss the center of the surface by.
missPupilX and missPupilY is how much we miss the center of the aperture by.
REAXp is true if we are aiming in x, false if we are aiming in y
        """     
        mfe = self.TheSystem.MFE
        op = mfe.AddOperand()
        if REAXp:
            op.ChangeType(constants.MeritOperandType_REAX)
        else:   
            op.ChangeType(constants.MeritOperandType_REAY)
        p1 = op.GetOperandCell(constants.MeritColumn_Param1)
        p1.IntegerValue = surf
        p3 = op.GetOperandCell(constants.MeritColumn_Param3)
        p3.DoubleValue = 0.0
        p4 = op.GetOperandCell(constants.MeritColumn_Param4)
        p4.DoubleValue = 0.0
        p5 = op.GetOperandCell(constants.MeritColumn_Param5)
        p5.DoubleValue = missPupilX
        p6 = op.GetOperandCell(constants.MeritColumn_Param6)
        p6.DoubleValue = missPupilY
        op.Target = missCenter
        wt = op.GetOperandCell(constants.MeritColumn_Weight)
        wt.DoubleValue = 1.0
        
    def AddREAOperands(self, surface, missx, missy, pupilx, pupily):
        """
        Add MF operands
        """
        self.AddREAOp(surface, missx, pupilx, pupily, True)
        self.AddREAOp(surface, missy, pupilx, pupily, False)

    def SurfaceDisplacement(self, surface, missx, missy):
        """ Displace the mirror vertex randomly """
        lde = self.TheSystem.LDE
        row = CastTo(lde,'IEditor').GetRowAt(surface - 1)
        colx = CastTo(row,'ILDERow').GetSurfaceCell(constants.SurfaceColumn_Par1);
        colx.DoubleValue = missx;
        coly = CastTo(row,'ILDERow').GetSurfaceCell(constants.SurfaceColumn_Par2);
        coly.DoubleValue = missy;
        
    def LocalOptimize(self, target):
        """
        Start local optimization, keep tunning until MF is below target, local optimization converges, 
        or 500 minutes has passed.
        """
        lopt = self.TheSystem.Tools.OpenLocalOptimization()
        lopt.Algorithm = constants.OptimizationAlgorithm_DampedLeastSquares
        lopt.Cycles = constants.OptimizationCycles_Infinite
        lopt.NumberOfCores = 8
        print("Starting local optimization")    
        CastTo(lopt, "ISystemTool").Run()
        mf = lopt.InitialMeritFunction
        counter = 0
        print("Starting loop, mf = " + str(mf))
        while mf > target:
            time.sleep(6)
            mf = lopt.CurrentMeritFunction
            print("mf = " + str(mf))
            counter = counter + 1
            if( counter > 10): break
        CastTo(lopt, "ISystemTool").Cancel()
        CastTo(lopt, "ISystemTool").Close()
        return(mf)

    def ListMirrorPlanes(self):
        """ Get a list containing the indexes of mirror surfaces """
        lde = self.TheSystem.LDE
        nSurf = lde.NumberOfSurfaces
        surfList = []
        for n in range(0,nSurf):
            surf = lde.GetSurfaceAt(n)
            if surf.Material == 'MIRROR':
                surfList.append(n)
        return(surfList)

    def createPickupsAndSetOrder(self, indexFrom, indexTo):
        """ Create picups with scale factor -1 for decenters and tilts. Set order to 1"""
        lde = self.TheSystem.LDE
        surf2 = lde.GetSurfaceAt(indexTo)
        for cellIndex in [12,13,14,15]:
            cell2 = CastTo(surf2, "IEditorRow").GetCellAt(cellIndex)
            pickup = cell2.CreateSolveType(constants.SolveType_SurfacePickup)._S_SurfacePickup
            pickup.ScaleFactor = -1.0
            pickup.Surface = indexFrom
            cell2.SetSolveData(pickup)
        ocol = CastTo(surf2,'ILDERow').GetSurfaceCell(constants.SurfaceColumn_Par6)
        ocol.IntegerValue = 1
        

    def CBify(self, index, variablep):
        """ Make surface a CG, make tilts variable """
        surf = self.TheSystem.LDE.GetSurfaceAt(index)
        setting = surf.GetSurfaceTypeSettings(constants.SurfaceType_CoordinateBreak)
        surf.ChangeType(setting)
        if(variablep):
            CastTo(surf,'IEditorRow').GetCellAt(14).MakeSolveVariable()
            CastTo(surf,'IEditorRow').GetCellAt(15).MakeSolveVariable()            

    def AddCoordinateBreaks(self):
        """ Add coordinate break surfaces to the LDE, set variables and pickups """ 
        mList = self.ListMirrorPlanes()
        lde = self.TheSystem.LDE
        for index in mList[::-1]:
            lde.InsertNewSurfaceAt(index+1)
            self.CBify(index+1, False)
            lde.InsertNewSurfaceAt(index)
            self.CBify(index, True)
            self.createPickupsAndSetOrder(index,index+2)

    def RemoveAllVariables(self):
        """ Remove all the variables """
        rv = self.TheSystem.Tools.RemoveAllVariables()

       
    def MisalignSystem(self, t1, t2, t3):
        """ Misalign the system
    T1 is the s.t.d. of decentering of the mirror
    T2 is the s.t.d. of vertex displacemnt from center of mirror
    T3 is the s.t.d. of how much the laser misses the center of the mirror

    The vertex is displaced by t1 + t2. The chief ray should then miss by t3 - t2.
    Normal mirros are displaced by t1 + t2, then missed by t3 - t2.
    Final plane is missed by t1 + t3, since there is no vertex displacement, only decenter and 
    """
        stopSurf = self.TheSystem.LDE.StopSurface
        stopRad = self.TheSystem.LDE.GetSurfaceAt(stopSurf).SemiDiameter
        lastSurf = self.TheSystem.LDE.NumberOfSurfaces - 1
        print('Last surface is' + str(lastSurf))
        px2 = random.gauss(0, t2)
        py2 = random.gauss(0,t2)
        px3 = random.gauss(0,t3)
        py3 = random.gauss(0,t3)
        px = (px3 - px2)/stopRad
        py = (py3 - py2)/stopRad

        mList = self.ListMirrorPlanes()
        mList.append(lastSurf)
        for surf in mList:
            x1 = random.gauss(0,t1)
            y1 = random.gauss(0,t1)
            x2 = random.gauss(0,t2)
            y2 = random.gauss(0,t2)
            x3 = random.gauss(0,t3)
            y3 = random.gauss(0,t3)
            if surf == stopSurf:
                x2 = px2
                y2 = py2
             
            if not surf == lastSurf:
                self.SurfaceDisplacement(surf, x1 + x2, y1 + y2)

            if surf == lastSurf:
                self.AddREAOperands(surf, x1 + x3, y1 + y3, px, py)
            elif not surf == stopSurf:
                self.AddREAOperands(surf, x3 - x2, y3 - y2, px, py)
            
        self.LocalOptimize(0.00000001)        
          
if __name__ == '__main__':
    #Make sure paths are ok before running
    # I have to open a ZOSAPI instance for every turn, or else it fails eventually
    # This slows down the process a whole lot
    for i in range(0,100):
        zosapi = MisAlignmentGenerator()
        print("Misaligning system " + str(i))
        zosapi.OpenFile('c:\\Users\haavagj\\tmp2.zmx', False)
        zosapi.RemoveAllMtfRows()
        zosapi.RemoveAllVariables()
        surfList = zosapi.ListMirrorPlanes()
        print(surfList)
        zosapi.AddCoordinateBreaks()
        zosapi.MisalignSystem(0.25,0.25,0.25)
        zosapi.TheSystem.SaveAs('c:\\Users\haavagj\\MC-alignment' + str(i) + '.zmx')
        del zosapi
