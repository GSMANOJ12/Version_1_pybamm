from sympy import false
from Cell_Test_Plan_Generator import *

class Simulation_cls(CellTestPlanGenerator):
    def __init__(self):
        super().__init__()

    def simulation_function(path,ow,initt,ind,cell_type):
        import pybamm

        self.declare_variables()

        self.sim_type="pybamm"
        # print(path,table_index)
        self.load_configuration_file()

        self.procees_inputs_for_pybamm(path,ind,ow,initt,cell_type)

        self.get_parameter_limit_from_opw()

        self.read_tables_from_word(self.test_spec_file_path)

        self.source_table_index = self.get_source_table_index()

        self.main_table = self.tables_list_with_names[self.source_table_index][1]

        self.get_reference_table()
        
        #Remove all the white spaces in the tables (except comment section) - Pre-processing
        self.main_table = self.pre_process_document(self.tables_list_with_names[self.source_table_index][1])

        #Find the column parameter positions
        self.get_column_iterators_positions(len(self.main_table[1]))

        #Configure the test plan
        #Find the temperature values
        if (self.sim_type == "pybamm"):
            #Find the temperature values
            #Expand the steps in the main table (expanding the loops)
            self.main_table_expanded = self.expand_the_table_steps(self.main_table, r'NA')

            #Fill the unknown values
            self.populate_unkown_values( r'NA', r'NA')

            table1=self.main_table_expanded[:]

            #Integrating operating window
            if (self.integrate_operating_window_flag==1):
                self.apply_operating_window_limits(self.main_table_expanded)

        prev_temp=0

        def func_to_val(expr):
            d={"irpt":self.cellcapacity,"inom":self.cellcapacity,"vdyn,max":self.upper_voltage_limit_value,"vdyn,min":self.lower_voltage_limit_value,"ipulse":30,"imax,cont":20,}
            dk=list(d.keys())
            exp1=re.findall(r'(\d*\.?\d*)\s*([*/⋅⋅.])\s*([a-zA-Z_][a-zA-Z0-9_]*)|([a-zA-Z_][a-zA-Z0-9_]*)\s*([*/⋅⋅.])\s*(\d*\.?\d*)',expr.replace(" ",""))      #k[0] is unit 
            div,mul=0,0
            flag1=0
            val1=None
            if exp1:
                for k1 in exp1:
                    for k1l in k1:
                        if '/' == k1l:
                            # print("div")
                            div=1
                        elif ('⋅' in k1l or '*' in k1l):
                            # print("mul")
                            mul=1
                        elif k1l.lower() in dk:
                            val1=d[k1l.lower()]
                            # print(val1,k1l.lower(),"ddddddddddddddddddddddd")
                        else:
                            val2r = re.search(r'(\d*\.?\d*)', k1l)
                            val2=val2r.group(0)
                            # print(val2,k1l,")
            else:
                if expr.lower() in dk:
                    val1=d[expr.lower()]
                else:
                    val12 = re.search(r'(\d*\.?\d*)\s*([*/⋅⋅.])\s*([a-zA-Z_][a-zA-Z0-9_]*)|([a-zA-Z_][a-zA-Z0-9_]*)\s*([*/⋅⋅.])\s*(\d*\.?\d*)|',expr.lower())
                    # print(val1,"here")
                    if val12:
                       val1=val12[0]
            if div and val1:
                fin = (float(val1) / float(val2))
                fin1=fin
                # print(fin1)
            elif mul and val1:
                fin = (float(val1) * float(val2))
                fin1=fin 
                # print(fin1)                           
            else:
                fin1=val1
                if bool(val1) == false:
                    val1 = re.findall(r'(\d+(\.\d+)?)', expr)
                    if val1:
                       val1=val1[0][0]
                    fin1=val1
                # if val1.isdigit():
                #     fin1=float(val1)    
                # print(fin1,"--------------------------------->")
            return fin1

        def convert_table_to_pybamm(step):
            global prev_temp
            l=list()
            if "command" in step[1].lower():  
                return None
            
            for ik in range(0,len(step)):
                step[ik]=step[ik].replace(" ","")
            # return step

            if step[1].lower()=="settemperature":
                temp = re.findall(r"(-?\d+\.?\d*?)\s*?([A-Za-z°]+)", step[2])                       # st=f"Rest for 5 seconds, temperature={i[2]}"
                temp=temp[0]
                if "C" in temp[1]:
                    temp=temp[0]
                elif "K" in temp[1]:
                    tmp = int(temp[0]) - 273.15
                    temp=tmp
                st="Rest for 20 seconds"
                prev_temp=temp
                return st,temp
                    
            if step[1].lower()=="rest":
                timee=re.findall(r"(\d+\.?\d*?)\s*?([a-zA-Z ]+)",step[3], re.IGNORECASE)                                   #n="".join(n)
                timee=timee[0]
                if "m" in timee[1]:
                    form="minutes"
                elif "h" in timee[1]:
                    form="hours"
                else:
                    form="seconds"
                # else:
                #     print("time format has not valid")
                time=timee[0]+" "+form
                st=f"Rest for {time}"
                return st,prev_temp
            
            rate=""
            # d={"irpt":self.cellcapacity,"inom":self.cellcapacity,"vdyn,max":self.upper_voltage_limit_value,"Vdyn,min":self.lower_voltage_limit_value}
            
            if step[1].lower()=="charge":
                pattern_par = re.finditer(r'(I|V)[<>=]((?:.(?!\bI=|\bV=))+)', step[2])
                results_par = [(m.group(1), m.group(2).strip(',')) for m in pattern_par]
                ct1=0
                if "SOC" in step[2]:
                    d={100: 3.496826, 95: 3.336521, 90: 3.335038, 85: 3.344210, 80: 3.334213, 75: 3.334198, 70: 3.331086, 65: 3.333976, 60: 3.314450, 55: 3.300315, 50: 3.298641, 45: 3.296620, 40: 3.294483, 35: 3.288390, 30: 3.277930, 25: 3.262439, 20: 3.244488, 15: 3.228439, 10: 3.210465, 0: 2.593900}
                    pattern_par = re.findall(r'(soc)[<>=](\d+(\.\d+)?)', step[2].lower())
                    # print(pattern_par)
                var,expr=results_par[0]               #################################################### took only for one experssion because of SOC
                if var.lower()=="i":
                    rate="Charge at "
                    if "C" in expr:
                        rate+=f"{expr}"
                    else:
                        v=func_to_val(expr)
                        rate+=f"{v} A"
                if var.lower()=="v":
                    v=func_to_val(expr)
                    rate="Hold at "
                    rate+=f"{v} V"
                pattern_time=re.findall(r'(t|tpulse)[<>=](\d+(\.\d+)?s?|[a-zA-Z_][a-zA-Z0-9_]*)', step[3].lower())
                pattern_ex = re.finditer(r'(I|V)[<>=]((?:.(?!\bI=|\bV=))+)', step[3])
                results_ex = [(m.group(1), m.group(2).strip(',')) for m in pattern_ex]
                if "SOC" in step[3]:
                    d1={100: 3.496826, 95: 3.336521, 90: 3.335038, 85: 3.344210, 80: 3.334213, 75: 3.334198, 70: 3.331086, 65: 3.333976, 60: 3.314450, 55: 3.300315, 50: 3.298641, 45: 3.296620, 40: 3.294483, 35: 3.288390, 30: 3.277930, 25: 3.262439, 20: 3.244488, 15: 3.228439, 10: 3.210465, 0: 2.593900}
                    pattern_par1 = re.findall(r'(soc)[<>=](\d+(\.\d+)?)', step[3].lower())
                    val=d1[int(pattern_par1[0][1])]
                    rate+= f" until {val} V"
                    # print()
                if results_ex:
                    var,expr=results_ex[0]               #################################################### took only for one experssion because of SOC
                    # print(var,expr) 
                    if var.lower()=="i":
                        rate+=" until"
                        if "c" in expr.lower():
                            rate+=f"{expr}"
                        else:
                            v=func_to_val(expr)
                            rate+=f" {v} A"
                    if var.lower()=="v":
                        rate+=" until"
                        v=func_to_val(expr)
                        rate+=f" {v} V" 

                if pattern_time: 
                  var=pattern_time[0][0]
                  expr=pattern_time[0][1]
                  if var.lower()=="t" or "tpulse" in var.lower():
                    rate+=" for"
                    v=func_to_val(expr)
                    if v:
                        rate+=f" {v} seconds"
                    else:
                        if 's' in expr:
                            rate+=f" {v} seconds"
                return rate,prev_temp


            if step[1].lower()=="discharge":
                pattern_par = re.finditer(r'(I|V)[<>=]((?:.(?!\bI=|\bV=))+)', step[2])
                results_par = [(m.group(1), m.group(2).strip(',')) for m in pattern_par]
                if "SOC" in step[2]:
                    d={100: 3.496826, 95: 3.336521, 90: 3.335038, 85: 3.344210, 80: 3.334213, 75: 3.334198, 70: 3.331086, 65: 3.333976, 60: 3.314450, 55: 3.300315, 50: 3.298641, 45: 3.296620, 40: 3.294483, 35: 3.288390, 30: 3.277930, 25: 3.262439, 20: 3.244488, 15: 3.228439, 10: 3.210465, 0: 2.593900}
                    pattern_par = re.findall(r'(soc)[<>=](\d+(\.\d+)?)', step[2].lower())
                    # print(pattern_par,"lklkmlmlmlmkmkmlkmlkmlm")
                var,expr=results_par[0]               #################################################### took only for one experssion because of SOC
                # print(var,expr)
                
                if var.lower()=="i":
                    rate="Discharge at "
                    if "C" in expr:
                        rate+=f"{expr}"
                    else:
                        v=func_to_val(expr)
                        rate+=f"{v} A"
                        # print("----------------",rate)
                if var.lower()=="v":
                    v=func_to_val(expr)
                    rate="Hold at "
                    rate+=f"{v} V"
                
                pattern_ex = re.finditer(r'(I|V)[<>=]((?:.(?!\bI=|\bV=))+)', step[3])
                results_ex = [(m.group(1), m.group(2).strip(',')) for m in pattern_ex]
                if "SOC" in step[3]:
                    d1={100: 3.496826, 95: 3.336521, 90: 3.335038, 85: 3.344210, 80: 3.334213, 75: 3.334198, 70: 3.331086, 65: 3.333976, 60: 3.314450, 55: 3.300315, 50: 3.298641, 45: 3.296620, 40: 3.294483, 35: 3.288390, 30: 3.277930, 25: 3.262439, 20: 3.244488, 15: 3.228439, 10: 3.210465, 0: 2.593900}
                    pattern_par1 = re.findall(r'(soc)[<>=](\d+(\.\d+)?)', step[3].lower())
                    val=d1[int(pattern_par1[0][1])]
                    rate+= f" until {val} V"
                pattern_time=re.findall(r'(t)[<>=](\d+(\.\d+)?|[a-zA-Z_][a-zA-Z0-9_]*)', step[3].lower())
                              #################################################### took only for one experssion because of SOC
                if results_ex:
                    var,expr=results_ex[0] 
                    if var.lower()=="i":
                        rate+=" until" 
                        if "c" in expr.lower():
                            rate+=f"{expr}"
                        else:
                            v=func_to_val(expr)
                            rate+=f" {v}A"
                    if var.lower()=="v":
                        rate+=" until" 
                        v=func_to_val(expr)
                        rate+=f" {v}V"    
                if pattern_time: 
                  var=pattern_time[0][0]
                  expr=pattern_time[0][1]
                  if var.lower()=="t":
                    rate+=" for"
                    v=func_to_val(expr)
                    rate+=f" {v}seconds"  
      
                return rate,prev_temp
                # for var, expr in results:
                #     if ct1>=1:
                #         rate+=" or "
                #     print(f"{var} = {expr}")
                #     if var.lower()=="i":
                #         if expr.lower() in d:
                #             rate+=d[expr.lower()]+"A"
                #         else:
                #             rate+=expr
                #     if var.lower()=="v":
                #         if expr.lower() in d:
                #             rate+=str(d[expr.lower()])+"V"
                #         else:
                #             rate+=expr
                #     ct1+=1
                # print(rate,"okokkokok")
                # exit_cond=step[4].split("=")
                # print(exit_cond,"llllllllllllll")
                # final=step[1]+" "+rate+" "+exit_cond
            # print(step[1],".........")
                # else:
                #     rate=step[2].split("=")
                #     print(rate)
                # exit_cond=step[3].split("=")
                # print(exit_cond)

        # print("ooooooooooooooo")
        l_step=[]
        # table1.append(['', 'Charge', 'I=Ipulse', 't>tpulse', 'Adjust the SOC to SoCRPT, Set'])
        # table1.append(['', 'Discharge', 'I=Ipulse', 't>tpulse', 'Adjust the SOC to SoCRPT, Set'])
        # l_step.append([str("Discharge at 2 C until 2.4 V"),str(25)+'oC'])
        # l_step.insert(0, ["Rest for 30 seconds", "25oC"])
        # l_step.insert(1, ["Charge at C/3 until 3.8 V", "25oC"])
        # l_step.insert(2, ["Rest for 1800 seconds", "25oC"])

        import pybamm
        # pybamm.step.string("Rest for 60 minutes")
        for ii in table1:
            # print(ii,"------------------------------------------------->")
            st=convert_table_to_pybamm(ii)
            # print(st,ii)
            if "start" in ii[1].lower():
                print("----------",ii[1],'-------------')
            if "end" in ii[1].lower():
                print("----------",ii[1],'-------------')
                # l_step.append([str("Charge at 1 C until 2.8 V"),str(25)+'oC'])
            if st:
               if "hold" in str(st[0]).lower():
                   continue
            #    print(f'pybamm.step.string("{st[0]}")')
               l_step.append([str(st[0]),str(st[1])+'oC'])
            # # if st:
            #     print(st[0],st[1])
            #     l_step.append([st[0],st[1]])
            # else:
            #     print(ii,"............>>>>>>>>>>>>>>>>>>>")

        # def format_cycle_steps(l_step):
        #     """Format all steps in cycles to ensure proper PyBaMM syntax"""
        #     formatted_steps = []
            
        #     for step in l_step:
        #         original_step = step[0]
        #         temperature = step[1]
                
        #         # Handle different step types
        #         if "charge" in original_step.lower():
        #             if "for" not in original_step.lower() and "until" not in original_step.lower():
        #                 # Add missing time duration
        #                 current = original_step.split("at")[1].strip()
        #                 formatted_step = f"Charge at {current} for 25 seconds"
        #                 print(original_step, "------------------->", formatted_step )
        #             else:
        #                 formatted_step = original_step
                        
        #         elif "rest" in original_step.lower():
        #             if "for" not in original_step.lower():
        #                 # Add missing time duration
        #                 formatted_step = f"Rest for 60 seconds"
        #                 print(original_step, "------------------->", formatted_step )   
        #             else:
        #                 formatted_step = original_step
                        
        #         else:
        #             # Default case
        #             formatted_step = original_step
                    
        #         formatted_steps.append([formatted_step, temperature])
                
        #     return formatted_steps
        
        # l_step = format_cycle_steps(l_step)

        model = pybamm.lithium_ion.SPM({"thermal": "lumped"})
        # print("jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj")
        # st=model.print_parameter_info(by_submodel=True)
        # # model.print_parameter_info()
        # print(st)
        # l_step.insert(0, ["Rest for 30 seconds", "25oC"])
        # l_step.insert(1, ["Charge at C/3 until 3.8 V", "25oC"])  # Add initial charging

        param = pybamm.ParameterValues("Chen2020")
        param.update({"Initial SEI thickness [m]": 0.2})
        param.update({
            'Negative electrode thickness [m]': 85e-6,  # Adjust for large format
            'Positive electrode thickness [m]': 75e-6,  # Adjust for large format
            'Separator thickness [m]': 25e-6,
            "Initial concentration in negative electrode [mol.m-3]": 0.2 * param["Maximum concentration in negative electrode [mol.m-3]"],
            "Initial concentration in positive electrode [mol.m-3]": 0.8 * param["Maximum concentration in positive electrode [mol.m-3]"],
        })
        if "bravo" in cell_type.lower():
            param.update({
            "Number of electrodes connected in parallel to make a cell": 15,  # Adjust for large format
            })
        else:
            param.update({
            "Number of electrodes connected in parallel to make a cell": 25,  # Adjust for large format
            })

        print("------------------------------------------------>>>>",self.cellcapacity)

        param.update({
            'Nominal cell capacity [A.h]': self.cellcapacity,
            'Open-circuit voltage at 0% SOC [V]': 2.394,
            'Open-circuit voltage at 100% SOC [V]': 3.428,
            'Lower voltage cut-off [V]': 1.9,
            'Upper voltage cut-off [V]': 4.1,
            "Number of cells connected in series to make a battery": 1.0,
        })

        # initial_soc = 0.2
        # # # y0 = pybamm.initial_conditions_from_soc(model, param, initial_soc)
        # param.process_model(model)
        # geometry = model.default_geometry
        # param.process_geometry(geometry)
        # # # # # #
        # mesh = pybamm.Mesh(geometry, model.default_submesh_types, model.default_var_pts,)
        # # disc = pybamm.SpectralVolume()
        # # # # # # # Step 6: Apply spatial methods
        # disc = pybamm.Discretisation(mesh, model.default_spatial_methods)
        # disc.process_model(model,inplace=False)  # ✅ Now spatial variables arefv ready
        # model.check_well_determined(disc)

        solver = pybamm.CasadiSolver( mode="fast with events",  # Fast solver with safe options
        dt_max=10.0,          # Reduced max timestep
        atol=1e-8,            # Tighter absolute tolerance
        rtol=1e-8,            # Tighter relative tolerance
    )
                                        # # mode="safe",
                                        # dt_max=30.0,    # Larger timestep
                                        # atol=1e-6,      # More relaxed tolerance
                                        # rtol=1e-6
                                    # )/

        # solver = pybamm.IDAKLUSolver( #mode="safe",  # More stable but slower
        #                             #   dt_max=1.0,   # Limit maximum time step
        #                               atol=1e-8,    # Absolute tolerance
        #                               rtol=1e-8,     # Relative tolerance,
        #                                 # max_nonlinear_iterations=100,  # Maximum nonlinear iterations
        #                             #     max_time=100000,  # Maximum simulation time
                                    #     # return_solution_if_failed_early=True,  # Return solution even if it fails early
                                    #   return_solution_if_failed_early=True,  # Return solution even if it fails early   
    
                                    #  )  #return_solution_if_failed_early=True

        # pybamm.step.string("Rest for 5 seconds", temperature="25oC"),
        #     pybamm.step.string("Rest for 30 minutes"),
        #     pybamm.step.string("Discharge at 1.0 A until 2.0 V"),
        #     pybamm.step.string("Discharge at 2.0 V until 1.2 A"),
        #     pybamm.step.string("Rest for 30 minutes"),
        #     pybamm.step.string("Charge at 1.0 A until 3.8 V")

        print("l_step",len(l_step))
        experiment1 = pybamm.Experiment([
            pybamm.step.string(i[0],temperature=i[1]) for i in l_step
            ])
        
        import matplotlib.pyplot as plt
        # t_eval = np.linspace(0, 100000 , 1000)
        # pybamm.Simulation.set_initial_soc(model,20)
        sim1 = pybamm.Simulation(model, parameter_values=param,experiment=experiment1,solver=solver)
        solution = sim1.solve([0,80000])
        sim1.plot(["Terminal voltage [V]","Current [A]"])


        plt.figure(figsize=(10, 8))
        plt.plot(solution["Time [s]"].entries, solution["Battery voltage [V]"].entries, label="Voltage")
        plt.xlabel("Time (s)")
        plt.ylabel("Voltage (V)")
        plt.title("Voltage vs Time")
        plt.grid(True)
        plt.legend()
        plt.tight_layout()
        plt.show()

        # # solution = sim.solve([0, 3600])
        # plot = pybamm.QuickPlot(sim1, figsize=(14, 7))
        # plot.plot(0.5)
        # 
        
        # # # Extract variables from solution
        # t = solution["Time [s]"].entries
        # voltage = solution["Terminal voltage [V]"].entries
        # current = solution["Current [A]"].entries
        
        # # # Plot manually
        # plt.figure()
        # plt.plot(t, voltage, label="Voltage (V)")
        # plt.plot(t, current, label="Current (A)")
        # plt.xlabel("Time (s)")
        # plt.ylabel("Value")
        # plt.legend()
        # plt.show()
        # def get_specific_capacity_positive(solution):
            # current = solution["Current [A]"].entries
            # time = solution["Time [s]"].entries
            # dt = np.diff(time)
            # current_avg = 0.5 * (current[:-1] + current[1:])
            # capacity = np.abs(np.sum(current_avg * dt)) / 3600  # Convert to Ah
            # return capacity
        # def run_cycle_simulation(model, parameter_values, num_cycles):
                # capacities = []
                # # Define one charge + discharge cycle
                # cycle = [
                #     "Discharge at 1C until 2.5V",
                #     "Charge at 1C until 4.2V"
                # ]
                # experiment_steps = cycle * num_cycles
                # experiment = pybamm.Experiment(experiment_steps)
                # sim = pybamm.Simulation(model, parameter_values=parameter_values, experiment=experiment)
                # solution = sim.solve()
                # # Get step-wise solutions directly
                # split_solutions = solution.cycles
                # for i in range(0, len(split_solutions), 2):  # Only Discharge steps
                #     discharge_solution = split_solutions[i]
                #     capacity = get_specific_capacity_positive(discharge_solution)
                #     capacities.append(capacity)
                # return capacities
        # model = pybamm.lithium_ion.DFN()  # You can also use pybamm.lithium_ion.Chen2020()
        # param = model.default_parameter_values
        # num_cycles = 5

        # capacities = run_cycle_simulation(model, param, num_cycles)

        # # Plotting
        # plt.figure(figsize=(8, 5))
        # plt.plot(range(1, num_cycles + 1), capacities, marker='o')
        # plt.xlabel("Cycle Number")
        # plt.ylabel("Specific Capacity (Positive) [Ah]")
        # plt.title("Specific Capacity (Positive) vs Number of Cycles")
        # plt.grid(True)
        # plt.tight_layout()
        # plt.show()


# def main():
#     Simulation_cls()

# if __name__ == '__main__':
#     main()


# # #a = r"C:\Users\INIYANK\OneDrive - Daimler Truck\Windows 10 - Backup\Folder@Work\Work - Folders\BMS Projects\Cell Specifications\Python Code\Cell Test Plan Generator\Test Specifications\DTC-P-10-1_Validation_tests_V1.1_changed.docx"
# # #a = r"C:\Users\INIYANK\OneDrive - Daimler Truck\Windows 10 - Backup\Folder@Work\Work - Folders\BMS Projects\Cell Specifications\Python Code\Cell Test Plan Generator\Test Specifications\DTC-P-6-1_Internal Resistance_V1.1_changed.docx"
# # #a = r"C:\Users\INIYANK\OneDrive - Daimler Truck\Windows 10 - Backup\Folder@Work\Work - Folders\BMS Projects\Cell Specifications\Python Code\Cell Test Plan Generator\Test Specifications\DTC-P-6-2_Internal Resistances BCC_V1.1_draft.docx"
# # a = r"C:\Users\INIYANK\OneDrive - Daimler Truck\Windows 10 - Backup\Folder@Work\Work - Folders\BMS Projects\Cell Specifications\Python Code\Cell Test Plan Generator\Test Specifications\DTC-P-2-2_Capacity_Energy_Efficiency_Temp_V1.0.docx"
# # "C:\Users\GSMANOJ\Downloads\Cell Test Plan Generator 3\Cell Test Plan Generator\Test Specifications\DTC-P-10-1_Validation_tests_V1.1_changed.docx"
# # C:\Users\GSMANOJ\Downloads\Cell Test Plan Generator_Copy\Cell Test Plan Generator\Supporting Documents\Operating Window_HDMD_CATL_B-sample_v2.8.xlsx
# # C:\Users\GSMANOJ\Downloads\Cell Test Plan Generator_Copy\Cell Test Plan Generator\Test Specifications\DTC-P-2-2_Capacity_Energy_Efficiency_Temp_V1.0.docx
# # C:\Users\GSMANOJ\Downloads\new\Cell Test Plan Generator_Copy\Cell Test Plan Generator\Test Specifications\DTC-P-2-1_Capacity_Energy_Efficiency_Crates_V1.0.docx
# a = r"C:\Users\GSMANOJ\Downloads\Cell Test Plan Generator_Copy\Cell Test Plan Generator\Test Specifications\DTC-P-2-1_Capacity_Energy_Efficiency_Crates_V1.0.docx"
# b = r"C:\Users\GSMANOJ\Downloads\Cell Test Plan Generator_Copy\Cell Test Plan Generator\Supporting Documents\Operating Window_HDMD_CATL_B-sample_v2.8.xlsx"

# c = r"C:\Users\GSMANOJ\Downloads\Cell Test Plan Generator_Copy\Cell Test Plan Generator\Supporting Documents\Initialization_Parameters_Document.xlsx"

# #d = 'Table 3 : Test procedure for thermal relaxation'
# #d = 'Table 2: : Test procedure for determination of internal resistance'
# # d = 'Table 5: Test procedure for Full DoD continuous current experiment'
# # d = 'Table 9: Test procedure for Charge neutral pulse experiment'
# # d = 'Table 2: Test procedure for determination of CC/3,y, EC/3,y at different temperatures'
# # d="Table 7: Test procedure for Partial DoD continuous current experiment"
# # Test procedure for determination of capacity CRPT, CC/x,25°C, EC/x 25°C and ηC/x,25°C at different discharging C-rates.
# d="Table 2: Test procedure for determination of capacity CRPT, CC/x,25°C, EC/x 25°C and ηC/x,25°C at different discharging C-rates."

# e = 'Cell Alpha - Op-Window'
# # e = "Cell Bravo - Op-Window"

# f = 'BaSyTec'

# g = 0

# h = 0

# i = ['Ah-Set', 'Ah-Charge', 'Ah-Discharge']


# # CellTestPlanGenerator.execute_main_generate_function(a, b, c, d, e, f, g, h, i)

# Simulation_cls.simulation_function(a,d,"pybamm",b,c,e)
