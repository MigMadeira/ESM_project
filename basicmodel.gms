*-----------------------------------------------------------------------------------------
* #Assignment 2021 - Course 42002#
*-----------------------------------------------------------------------------------------
* * Group:              4.2
* * Author(s):          Miguel Madeira
*-----------------------------------------------------------------------------------------

* Input from Excel File
*-----------------------------------------------------------------------------------------
*sets
$call =xls2gms   r="Timesteps"!C3:C8762          i=Data_GroupWork.xlsx    o=set_tt.inc
$call =xls2gms   r="Timesteps"!D2:D8762          i=Data_GroupWork.xlsx    o=set_tt_0.inc

*electricity consumption data
$call =xls2gms   r="ElectConsumption"!B6:C8765   i=Data_GroupWork.xlsx    o=par_ElectCons_P4_3_A80.inc
$call =xls2gms   r="ElectConsumption"!D6:E8765   i=Data_GroupWork.xlsx    o=par_ElectCons_P4_3_A180.inc
$call =xls2gms   r="ElectConsumption"!F6:G8765   i=Data_GroupWork.xlsx    o=par_ElectCons_P4_1_A80.inc
$call =xls2gms   r="ElectConsumption"!H6:I8765   i=Data_GroupWork.xlsx    o=par_ElectCons_P4_1_A180.inc

*heat consumption data
$call =xls2gms   r="HeatConsumption"!B4:C8763   i=Data_GroupWork.xlsx    o=par_HeatCons_A80.inc
$call =xls2gms   r="HeatConsumption"!D4:E8763   i=Data_GroupWork.xlsx    o=par_HeatCons_A180.inc

*spot price data
$call =xls2gms   r="ElectSpotPrices"!D2:E8761   i=Data_GroupWork.xlsx    o=par_ElectSpot.inc


*-----------------------------------------------------------------------------------------
* Set declaration and definition
*-----------------------------------------------------------------------------------------
set tt_0   timesteps (including tt0)
/
$include set_tt_0.inc
/
;

set tt(tt_0)  timesteps (without tt0)
/
$include set_tt.inc
/
;
*-----------------------------------------------------------------------------------------

*-----------------------------------------------------------------------------------------
* Parameter declaration and definition
*-----------------------------------------------------------------------------------------

parameter elect_load_P4_3_A80(tt) electricity demand per timestep [kWh_el]
/
$include par_ElectCons_P4_3_A80.inc
/
;

parameter elect_load_P4_3_A180(tt) electricity demand per timestep [kWh_el]
/
$include par_ElectCons_P4_3_A180.inc
/
;

parameter elect_load_P4_1_A80(tt) electricity demand per timestep [kWh_el]
/
$include par_ElectCons_P4_1_A80.inc
/
;

parameter elect_load_P4_1_A180(tt) electricity demand per timestep [kWh_el]
/
$include par_ElectCons_P4_1_A180.inc
/
;

parameter heat_load_A80(tt) electricity demand per timestep [kWh_heat]
/
$include par_HeatCons_A80.inc
/
;


parameter heat_load_A180(tt) electricity demand per timestep [kWh_heat]
/
$include par_HeatCons_A180.inc
/
;

parameter spot_price(tt) spot market price per timestep [€ * (kWh_el)^(-1)]
/
$include par_ElectSpot.inc
/
;
spot_price(tt) = spot_price(tt)/1000;
*-----------------------------------------------------------------------------------------

*-----------------------------------------------------------------------------------------
* Scalar declaration and definition
*-----------------------------------------------------------------------------------------

Scalar
         time_step           smallest time step [h] /1/
         num_weeks           number of weeks per year considered /9/

         cost_gas            cost of natural gas [€ * (kWh)^(-1)] /0.09/
         
         boiler_CC           boiler capital cost [€ * (kW)^(-1)] /63.83/
         lifetime_boiler     boiler's lifetime (years) /25/
         OaM_cost_boiler     Operation and maintenance cost boiler[€ * (kW)^(-1)] /0.0011/

         taxes_electCons     taxes electricity consumption [€ * (kWh)^(-1)] /0.12/
         taxes_electHeat     taxes electricity for heating [€ * (kWh)^(-1)] /0.036/
         taxes_gasCons       taxes gas consumption [€ * (kWh)^(-1)] /0.04/
         fees_electNetwork   network costs for electricity consumption [€ * (kWh)^(-1)] /0.15/
;         

*-----------------------------------------------------------------------------------------
* Variables declaration
*-----------------------------------------------------------------------------------------

Variable
         Z            energy purchase cost [€]
         T_grid       total anual grid costs
         T_boiler     total anual boiler costs
;

Positive Variable
         x_el_grid(tt) electricity from grid [kWh_el]
         x_th_boil(tt) output of heat boiler [kWh_th]
;

Integer Variable
         boilerC       installed boiler capacity in kW
;

*-----------------------------------------------------------------------------------------
* Equations declaration
*-----------------------------------------------------------------------------------------

Equations
obj                           Objective function
grid_cost                     Total grid cost calculation
boiler_cost                   Total boiler cost calculation

energycons_P4_3_A80(tt)       Energy bought must be higher than the consumption _P4_3_A80
energycons_P4_3_A180(tt)      Energy bought must be higher than the consumption _P4_3_A180
energycons_P4_1_A80(tt)       Energy bought must be higher than the consumption _P4_1_A80
energycons_P4_1_A180(tt)      Energy bought must be higher than the consumption _P4_1_A180

heatcons_A80(tt)              Heat bought or generated must be higher than the consumption _A80
heatcons_A180(tt)             Heat bought or generated must be higher than the consumption _A180

boiler_install_capacity(tt)       Output of heat boiler must be less than it's capacity
;

obj ..   Z =e= T_grid + T_boiler;
                
grid_cost        .. T_grid      =e= sum(tt, spot_price(tt)*x_el_grid(tt) + taxes_electCons*x_el_grid(tt)+fees_electNetwork*x_el_grid(tt));
boiler_cost      .. T_boiler    =e= sum(tt,cost_gas*x_th_boil(tt) + taxes_gasCons*x_th_boil(tt))+ boiler_CC*boilerC/lifetime_boiler + OaM_cost_boiler*boilerC;

energycons_P4_3_A80(tt)  .. x_el_grid(tt) =e= elect_load_P4_3_A80(tt);
energycons_P4_3_A180(tt) .. x_el_grid(tt) =e= elect_load_P4_3_A180(tt);
energycons_P4_1_A80(tt)  .. x_el_grid(tt) =e= elect_load_P4_1_A80(tt);
energycons_P4_1_A180(tt) .. x_el_grid(tt) =e= elect_load_P4_1_A180(tt);

heatcons_A80(tt)  .. x_th_boil(tt) =e= heat_load_A80(tt);
heatcons_A180(tt) .. x_th_boil(tt) =e= heat_load_A180(tt);

boiler_install_capacity(tt) .. x_th_boil(tt) =l= boilerC;


Model energy_P4_3_A80 /obj, energycons_P4_3_A80, heatcons_A80, boiler_install_capacity, grid_cost, boiler_cost/ ;
option mip=cplex;
solve energy_P4_3_A80 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;

*execute_unload "excelfiles/basicresultsP4_3_A80.gdx" Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;
*
*
*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/basicresultsP4_3_A80.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)  
*writer.save()   
*endEmbeddedCode

Model energy_P4_3_A180 /obj, energycons_P4_3_A180, heatcons_A180, boiler_install_capacity, grid_cost, boiler_cost/ ;
option mip=cplex;
solve energy_P4_3_A180 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;

*execute_unload "excelfiles/basicresultsP4_3_A180.gdx" Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;
*
*
*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/basicresultsP4_3_A180.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)  
*writer.save()   
*endEmbeddedCode

Model energy_P4_1_A80 /obj, energycons_P4_1_A80, heatcons_A80, boiler_install_capacity, grid_cost, boiler_cost/ ;
option mip=cplex;
solve energy_P4_1_A80 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;

*execute_unload "excelfiles/basicresultsP4_1_A80.gdx" Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;
*
*
*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/basicresultsP4_1_A80.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)  
*writer.save()   
*endEmbeddedCode

Model energy_P4_1_A180 /obj, energycons_P4_1_A180, heatcons_A180, boiler_install_capacity, grid_cost, boiler_cost/ ;
option mip=cplex;
solve energy_P4_1_A180 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;

*execute_unload "excelfiles/basicresultsP4_1_A180.gdx" Z.l, x_el_grid.l, x_th_boil.l,boilerC.l,T_grid.l,T_boiler.l;
*
*
*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/basicresultsP4_1_A180.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)  
*writer.save()   
*endEmbeddedCode