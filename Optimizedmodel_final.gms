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

*solar irradiance
$call =xls2gms   r="solar_data"!D2:E8761        i=solar_data.xlsx        o=par_solardata.inc
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


parameter irradiance(tt) irradiance per timestep [Wm^(-2)]
/
$include par_solardata.inc
/
;
irradiance(tt) = irradiance(tt)/1000;
*the irradiance is in watts in our data so we convert it to kW.
*-----------------------------------------------------------------------------------------


parameter soc_initial(tt_0) state of charge initial /tt0 0/;

*-----------------------------------------------------------------------------------------
* Scalar declaration and definition
*-----------------------------------------------------------------------------------------

Scalar
         time_step       smallest time step [h] /1/
*         num_weeks       number of weeks per year considered /9/

         cost_gas        cost of natural gas [€ * (kWh)^(-1)] /0.09/
        
*-----------------------------------------------------------------------------------------
* Boiler scalars
*-----------------------------------------------------------------------------------------
 
         boiler_CC           boiler capital cost [€ * (kW)^(-1)] /63.83/
         lifetime_boiler     boiler's lifetime (years) /25/
         OaM_cost_boiler     Operation and maintenance cost boiler[€ * (kWh)^(-1)] /0.0011/
*-----------------------------------------------------------------------------------------
* heat pump scalars
*-----------------------------------------------------------------------------------------     
         HP_CC               Heatpump capital cost [€ * (kW)^(-1)] /1402/
         HP_lifetime         Heatpump lifetime (years)  /25/
         OaM_cost_HP         Operation and maintenance cost HP [€ * (kWh)^(-1)]/0.0027/
         COP_HP              Efficiency of the heat pump /2.9/

*-----------------------------------------------------------------------------------------
* Taxes and fees
*-----------------------------------------------------------------------------------------     
         taxes_electCons     taxes electricity consumption [€ * (kWh)^(-1)] /0.12/
         taxes_electHeat     taxes electricity for heating [€ * (kWh)^(-1)] /0.036/
         taxes_gasCons       taxes gas consumption [€ * (kWh)^(-1)] /0.04/
         fees_electNetwork   network costs for electricity consumption [€ * (kWh)^(-1)] /0.015/
 
*-----------------------------------------------------------------------------------------
* Storage scalars
*-----------------------------------------------------------------------------------------        
         bat_max_C                max battery capacity(kWh) /55/
         stor_min_C               min battery and heat storage capacity /0/
         CC_bat                   battery capital costs (€ kWh^-1) /1073/
         lifetime_bat             battery lifetime(years) /20/
         bat_charge_eff           battery charging effficiency /0.98/
         bat_discharge_eff        battery discharging effficiency /0.97/
         bat_linear_loss          battery linear losses /0.00004167/
         bat_charge_lim           battery charge limit /1/  
         bat_discharge_lim        battery discharge limit /0.5/
         OaM_cost_bat             battery Operation and maintenance cost [€ * (kWh)^(-1)] /0.0021/
	
         heat_stor_max_C          max heat storage capacity(kWh)                         /55/
         CC_heat_stor             heat storage capital costs (€ kWh^-1)                  /422.1/
         lifetime_heat_stor       heat storage lifetime(years)                           /30/
         heat_stor_charge_eff     heat storage charging efficiency                       /1/
         heat_stor_discharge_eff  heat storage discharging efficiency                    /1/
         heat_stor_linear_loss    heat storage linear losses                             /0.021/
         heat_stor_charge_lim     heat storage charge limit                              /1/  
         heat_stor_discharge_lim  heat storage discharge limit                           /1/
         OaM_cost_heat_stor       battery Operation and maintenance cost [€ * (kWh)^(-1)]/0.0007/
         
*-----------------------------------------------------------------------------------------
* PV scalars
*-----------------------------------------------------------------------------------------        
         PV_max_C                max PV capacity(kWh) /20/
         CC_PV                   PV capital costs (€ kWh^-1) /1177/
         efficiency_PV           Efficiency of the solar cell /0.18/
         lifetime_PV             PV lifetime(years) /25/
         OaM_cost_PV             battery Operation and maintenance cost [€ * (kWh)^(-1)] /0/
;
*-----------------------------------------------------------------------------------------
* Variables declaration
*-----------------------------------------------------------------------------------------

Variable

         Z            energy purchase cost [€]
         T_grid       total anual grid costs
         T_boiler     total anual boiler costs
         T_HP         total anual heat pump costs
         T_PV         total anual solar cell costs
         T_bat        total anual battery costs
         T_heat_stor  total anual heat storage costs
         T_surplus    total anual earnings from selling electricity
;

Positive Variable
         x_el_grid(tt) electricity from grid        [kWh_el]
         x_el_PV(tt)   electricity from solar cells [kWh_el] going to electricity consumption
         PV_to_HP(tt)  electricity from solar cells [kWh_el] going to heat pump
         x_th_boil(tt) output of heat boiler        [kWh_th]
         x_th_HP(tt)   output of heat pump          [kWh_th]
         x_el_HP(tt)   electrical consumption heat pump         
         energy_surplus(tt) exported electricity to the grid [kWh_el]
*-----------------------------------------------------------------------------------------
* Storage variables
*-----------------------------------------------------------------------------------------                 
         soc_bat(tt_0)              state of charge battery
         charge_bat(tt_0)           charge battery
         discharge_bat(tt_0)        discharge battery
         bat_to_HP(tt_0)            battery electricity to heat pump

         soc_heat_stor(tt_0)        state of charge heat storage
         charge_heat_stor(tt_0)     charge heat storage
         discharge_heat_stor(tt_0)  discharge heat storage
         
         PVarea                     installed PV area

;

Integer Variable

         boilerC       installed boiler capacity in kW
         HPC           installed heat pump capacity in kW
         PVC           installed PV capacity in kW
      
         bat           installed battery capacity in kW
         heat_stor     installed heat storage unit capacity in kW
         
;

*bat.lo = 1;
bat.up = bat_max_C;
heat_stor.up = heat_stor_max_C;
PVC.up = PV_max_C;

*-----------------------------------------------------------------------------------------
* Equations declaration
*-----------------------------------------------------------------------------------------

Equations
obj                           Objective function
grid_cost                     Total grid cost calculation
boiler_cost                   Total boiler cost calculation
HP_cost                       Total heat pump cost calculation
PV_cost                       Total solar cell cost calculation
bat_cost                      Total battery cost calculation
heat_stor_cost                Total heat storage cost calculation
surplus_earnings              Total surplus earnings

*-----------------------------------------------------------------------------------------
* Equilibrium equation declaration
*-----------------------------------------------------------------------------------------
energycons_P4_3_A80(tt)       Energy bought must be higher than the consumption _P4_3_A80
energycons_P4_3_A180(tt)      Energy bought must be higher than the consumption _P4_3_A180
energycons_P4_1_A80(tt)       Energy bought must be higher than the consumption _P4_1_A80
energycons_P4_1_A180(tt)      Energy bought must be higher than the consumption _P4_1_A180

heatcons_A80(tt)              Heat bought or generated must be higher than the consumption _A80
heatcons_A180(tt)             Heat bought or generated must be higher than the consumption _A180

*-----------------------------------------------------------------------------------------
* Boiler and Heat Pump capacity equation declaration
*-----------------------------------------------------------------------------------------

boiler_install_capacity(tt)     Output of heat boiler must be less than it's capacity
HP_install_capacity(tt)         Output of HP must be less than it´s capacity
HP_el_consumption(tt)           heat pump electrical consumption calculation

*-----------------------------------------------------------------------------------------
* Storage Equation declaration
*note: I assumed the battery and storage is bought at the start of the year if bought at all (tt_0)
*-----------------------------------------------------------------------------------------     
bat_SOC_ini(tt_0)               battery soc initialisation 
bat_SOC_const(tt_0)             battery soc calculation or constraint
bat_SOC_max(tt_0)               battery soc must be less than the maximum capacity
bat_charge_max_1dt(tt_0)        battery can only charge some amount on a time dt
bat_discharge_max_1dt(tt_0)     battery can only discharge some amount on a time dt
             


heat_stor_SOC_ini(tt_0)         heat storage soc initialisation 
heat_stor_SOC_const(tt_0)       heat storage soc calculation or constraint
heat_stor_SOC_max(tt)           heat storage soc must be less than the maximum capacity
heat_stor_charge_max_1dt(tt)    heat storage can only charge some amount on a time dt
heat_stor_discharge_max_1dt(tt) heat storage can only charge some amount on a time dt

*-----------------------------------------------------------------------------------------
* PV Equation declaration
*-----------------------------------------------------------------------------------------

PV_production(tt)                 electricity produced by the PV on a time dt
PV_capacity(tt)                   electricity produced by the PV on a time dt
PV_max_area_80                    PV max installed m^2 based on roof size for the 80 m^2 house
PV_max_area_180                   PV max installed m^2 based on roof size for the 180 m^2 house
;


*-----------------------------------------------------------------------------------------
* Objective function definition
*-----------------------------------------------------------------------------------------  

obj              .. Z =e= T_grid + T_boiler + T_HP + T_bat + T_heat_stor + T_PV - T_surplus;
                 
grid_cost        .. T_grid      =e= sum(tt, spot_price(tt)*x_el_grid(tt) + taxes_electCons*x_el_grid(tt)+fees_electNetwork*x_el_grid(tt));
boiler_cost      .. T_boiler    =e= sum(tt,cost_gas*x_th_boil(tt) + taxes_gasCons*x_th_boil(tt))+ boiler_CC*boilerC/lifetime_boiler + OaM_cost_boiler*boilerC;
HP_cost          .. T_HP        =e= sum(tt,x_el_HP(tt)*spot_price(tt) + taxes_electHeat*x_el_HP(tt)+ fees_electNetwork*x_el_HP(tt)) + HP_CC*HPC/HP_lifetime + OaM_cost_HP*HPC;
PV_cost          .. T_PV        =e= CC_PV*PVC/lifetime_PV;
bat_cost         .. T_bat       =e= CC_bat*bat/lifetime_bat + OaM_cost_bat*bat;
heat_stor_cost   .. T_heat_stor =e= CC_heat_stor*heat_stor/lifetime_heat_stor + OaM_cost_heat_stor*heat_stor;
surplus_earnings .. T_surplus   =e= sum(tt,spot_price(tt)*energy_surplus(tt));

*-----------------------------------------------------------------------------------------
* Equilibrium equation definiton
*-----------------------------------------------------------------------------------------
energycons_P4_3_A80(tt)  .. x_el_grid(tt) + discharge_bat(tt) + x_el_PV(tt) - charge_bat(tt) - energy_surplus(tt) =g= elect_load_P4_3_A80(tt)  ;
energycons_P4_3_A180(tt) .. x_el_grid(tt) + discharge_bat(tt) + x_el_PV(tt) - charge_bat(tt) - energy_surplus(tt) =g= elect_load_P4_3_A180(tt) ;
energycons_P4_1_A80(tt)  .. x_el_grid(tt) + discharge_bat(tt) + x_el_PV(tt) - charge_bat(tt) - energy_surplus(tt) =g= elect_load_P4_1_A80(tt)  ;
energycons_P4_1_A180(tt) .. x_el_grid(tt) + discharge_bat(tt) + x_el_PV(tt) - charge_bat(tt) - energy_surplus(tt) =g= elect_load_P4_1_A180(tt) ;

heatcons_A80(tt)  .. x_th_boil(tt) + x_th_HP(tt)  + discharge_heat_stor(tt) - charge_heat_stor(tt) =g= heat_load_A80(tt);
heatcons_A180(tt) .. x_th_boil(tt) + x_th_HP(tt)  + discharge_heat_stor(tt) - charge_heat_stor(tt) =g= heat_load_A180(tt);

*-----------------------------------------------------------------------------------------
* Boiler and Heat Pump capacity equation definition
*-----------------------------------------------------------------------------------------

boiler_install_capacity(tt) .. x_th_boil(tt) =l= boilerC;
HP_install_capacity(tt)     .. x_th_HP(tt)   =l= HPC;
HP_el_consumption(tt)       .. COP_HP*(x_el_HP(tt) + PV_to_HP(tt) + bat_to_HP(tt)) =e= x_th_HP(tt);
*-----------------------------------------------------------------------------------------
* Storage Equation definition
*-----------------------------------------------------------------------------------------      

bat_SOC_ini(tt_0)$(ord(tt_0)=1) .. soc_bat(tt_0) =e= soc_initial(tt_0);
bat_SOC_const(tt_0)$(tt(tt_0))  .. soc_bat(tt_0) =e= (1 - bat_linear_loss)*soc_bat(tt_0-1) + charge_bat(tt_0) * bat_charge_eff
                                                        - discharge_bat(tt_0)/bat_discharge_eff - bat_to_HP(tt_0)/bat_discharge_eff;
bat_SOC_max(tt)                 .. soc_bat(tt) =l= bat;
bat_charge_max_1dt(tt)          .. charge_bat(tt) =l= bat*bat_charge_lim;
bat_discharge_max_1dt(tt)       .. discharge_bat(tt) + bat_to_HP(tt) =l= bat_discharge_lim*bat;



heat_stor_SOC_ini(tt_0)$(ord(tt_0)=1)  .. soc_heat_stor(tt_0) =e= soc_initial(tt_0);
heat_stor_SOC_const(tt_0)$(tt(tt_0))   .. soc_heat_stor(tt_0) =e= soc_heat_stor(tt_0-1) + charge_heat_stor(tt_0)*heat_stor_charge_eff
                                                                - discharge_heat_stor(tt_0)/heat_stor_discharge_eff - heat_stor_linear_loss*soc_heat_stor(tt_0);
heat_stor_SOC_max(tt)                  .. soc_heat_stor(tt) =l= heat_stor;
heat_stor_charge_max_1dt(tt)           .. charge_heat_stor(tt) =l= heat_stor_charge_lim*heat_stor;
heat_stor_discharge_max_1dt(tt)        .. discharge_heat_stor(tt) =l= heat_stor_discharge_lim*heat_stor;

*-----------------------------------------------------------------------------------------
* PV Equation definition
*-----------------------------------------------------------------------------------------
      
PV_production(tt)                   .. x_el_PV(tt) + PV_to_HP(tt) =e= irradiance(tt)*PVarea*efficiency_PV;
PV_capacity  (tt)                   .. x_el_PV(tt) + PV_to_HP(tt) =l= PVC;
PV_max_area_80                      .. PVarea =l= 80*0.8;
PV_max_area_180                     .. PVarea =l= 180*0.8;
*-----------------------------------------------------------------------------------------
* Models
*-----------------------------------------------------------------------------------------      

Model energy_P4_3_A80 / obj, energycons_P4_3_A80, heatcons_A80, boiler_install_capacity, HP_install_capacity,HP_el_consumption,
                        bat_SOC_ini, bat_SOC_const, bat_SOC_max, bat_charge_max_1dt, bat_discharge_max_1dt, 
                        heat_stor_SOC_ini,heat_stor_SOC_const,heat_stor_SOC_max,heat_stor_charge_max_1dt,heat_stor_discharge_max_1dt,
                        grid_cost,boiler_cost,HP_cost,bat_cost,heat_stor_cost,PV_cost,PV_production,PV_max_area_80,PV_capacity,surplus_earnings/ ;
option mip=cplex;
solve energy_P4_3_A80 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,x_el_HP.l,
            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l,bat_to_hp.l,PV_to_HP.l;

*This commented code is used to extract the results either to an excel or to a .gdx file
*execute_unload "excelfiles/resultsP4_3_A80.gdx" Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,
*            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
*            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l;
**


*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/resultsP4_3_A80.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)  
*writer.save()   
*endEmbeddedCode

Model energy_P4_3_A180 /obj, energycons_P4_3_A180, heatcons_A180, boiler_install_capacity, HP_install_capacity,HP_el_consumption,
                        bat_SOC_ini, bat_SOC_const, bat_SOC_max, bat_charge_max_1dt, bat_discharge_max_1dt, 
                        heat_stor_SOC_ini,heat_stor_SOC_const,heat_stor_SOC_max,heat_stor_charge_max_1dt,heat_stor_discharge_max_1dt,
                        grid_cost,boiler_cost,HP_cost,bat_cost,heat_stor_cost,PV_cost,PV_production,PV_max_area_180,PV_capacity,surplus_earnings/ ;
option mip=cplex;
solve energy_P4_3_A180 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,x_el_HP.l,
            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l,bat_to_hp.l,PV_to_HP.l;

*This commented code is used to extract the results either to an excel or to a .gdx file
*execute_unload "excelfiles/resultsP4_3_A180.gdx" Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,
*            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
*            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l;
*            

*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/resultsP4_3_A180.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)
*writer.save()   
*endEmbeddedCode
           
Model energy_P4_1_A80 /obj, energycons_P4_1_A80, heatcons_A80, boiler_install_capacity, HP_install_capacity,HP_el_consumption,
                        bat_SOC_ini, bat_SOC_const, bat_SOC_max, bat_charge_max_1dt, bat_discharge_max_1dt, 
                        heat_stor_SOC_ini,heat_stor_SOC_const,heat_stor_SOC_max,heat_stor_charge_max_1dt,heat_stor_discharge_max_1dt,
                        grid_cost,boiler_cost,HP_cost,bat_cost,heat_stor_cost,PV_cost,PV_production,PV_max_area_80,PV_capacity,surplus_earnings/ ;
option mip=cplex;
solve energy_P4_1_A80 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,x_el_HP.l,
            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l,bat_to_hp.l,PV_to_HP.l;

*This commented code is used to extract the results either to an excel or to a .gdx file
*execute_unload "excelfiles/resultsP4_1_A80.gdx" Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,
*            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
*            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l;
*            
*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/resultsP4_1_A80.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)
*writer.save()   
*endEmbeddedCode

Model energy_P4_1_A180 /obj, energycons_P4_1_A180, heatcons_A180, boiler_install_capacity, HP_install_capacity,HP_el_consumption,
                        bat_SOC_ini, bat_SOC_const, bat_SOC_max, bat_charge_max_1dt, bat_discharge_max_1dt, 
                        heat_stor_SOC_ini,heat_stor_SOC_const,heat_stor_SOC_max,heat_stor_charge_max_1dt,heat_stor_discharge_max_1dt,
                        grid_cost,boiler_cost,HP_cost,bat_cost,heat_stor_cost,PV_cost,PV_production,PV_max_area_180,PV_capacity,surplus_earnings/ ;
option mip=cplex;
solve energy_P4_1_A180 using mip minimizing Z;
Display Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,x_el_HP.l,
            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l,bat_to_hp.l,PV_to_HP.l;


*This commented code is used to extract the results either to an excel or to a .gdx file
*execute_unload "excelfiles/resultsP4_1_A180.gdx" Z.l, x_el_grid.l, x_th_boil.l,Soc_bat.l,bat.l,heat_stor.l,boilerC.l, x_th_HP.l,HPC.l,
*            charge_heat_stor.l,discharge_heat_stor.l,T_grid.l,T_boiler.l,T_HP.l,T_bat.l,T_heat_stor.l,
*            T_PV.l,T_surplus.l,PVC.l,PVarea.l,x_el_PV.l,energy_surplus.l;
*            

*embeddedCode Python:
*import pandas as pd
*import gams2numpy
*gams.wsWorkingDir = 'E:/MEFT/erasmus/1º semestre/Energy system modelling/Project/Final code'
*g2np = gams2numpy.Gams2Numpy(gams.ws.system_directory)
*writer = pd.ExcelWriter('excelfiles/resultsP4_1_A180.xlsx', engine='openpyxl')
*for sym in gams.db:
*   arr = g2np.gmdReadSymbolStr(gams.db, sym.name)
*   pd.DataFrame(data=arr).to_excel(writer, sheet_name=sym.name)
*writer.save()   
*endEmbeddedCode