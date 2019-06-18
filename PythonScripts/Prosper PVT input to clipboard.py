# Objective: Copy PVT data from a Prosper file and copy a formatted view to clipboard ready for Excel.
# Author: Thorjan Knudsvik
from PetexOpenServer import *
import os

DoCmd('PROSPER.MENU.UNITS.NORSI')
DoSet('PROSPER.PVT.Calc.Mode', 1)
DoSet('PROSPER.PVT.Calc.TempUser[0]', 15)
DoSet('PROSPER.PVT.Calc.PresUser[0]', 1)
DoCmd('PROSPER.PVT.CALC')
T = DoGet('PROSPER.PVT.Calc.Results[0].Temp')
P = DoGet('PROSPER.PVT.Calc.Results[0].Pres')
Z = round(DoGet('PROSPER.PVT.Calc.Results[0].ZFactor'), 3)
rho_w = round(DoGet('PROSPER.PVT.Calc.Results[0].WatDen'), 1)
MW = round(DoGet('PROSPER.PVT.Input.Grvgas') * 28.97, 1)
mu_o  = DoGet('PROSPER.PVT.Calc.Results[0].OilVis') / 1000
mu_w = DoGet('PROSPER.PVT.Calc.Results[0].WatVis') / 1000
mu_g = DoGet('PROSPER.PVT.Calc.Results[0].GasVis') / 1000
if DoGet('PROSPER.SIN.SUM.Fluid') == 0:
	rho_o = round(DoGet('PROSPER.PVT.Calc.Results[0].OilDen'), 1)
elif DoGet('PROSPER.SIN.SUM.Fluid') == 1:
	print('Gas condensate model')
	quit()
else:
	print('Model type is not defined.')
	quit()

a = str(Z) + '\t' + str(rho_o) + '\t' + str(rho_w) + '\t' + str(MW) + '\t' + '{:0.2e}'.format(
	mu_o) + '\t' + '{:0.2e}'.format(mu_w) + '\t' + '{:0.2e}'.format(mu_g)

print('Temperature:', T, ', Pressure:', P)
print('Input: ', a)
os.system("echo {} | clip".format(a))
DoCmd('PROSPER.SHUTDOWN')
print('Excel input copied to clipboard')