    
from graphviz import *
def flowchart_function(table,flowchart_output,output_folder):
        # from graphviz import * self.flowchart_function(self.main_table)

        # tables = self.read_tables_from_word(path)
        # print(tables)
        # source_table_index = self.source_table_index(tables)
        # print(source_table_index)
        # tables[source_table_index] = self.pre_process_document(tables[source_table_index])



        func = []
        parameter = []
        exit_condition = []
        print(func, "\n", len(func))
        for i in table:
            func.append(i[1])
            parameter.append(i[2])
            exit_condition.append(i[3])
        dot = Digraph()
        dot.attr(rankdir='LR')                                # LR: Left to Right
        dot.attr(rank='same')                               # dot.attr(size='5,20')                               # Size in inches (Width x Height)
        # dot.attr(ratio='compress')                        # Try to fit the content in the size
        dot.attr(splines='ortho')                          # Straighter lines
        dot.attr(nodesep='0.4')                            # default is around 0.25
        dot.attr(ranksep='1.0')                            # default is around 0.5
        dot.attr(concentrate='false')
        inc1=0
        for i in range(1, len(func)):                         # creating the nodes or block one by one
            with dot.subgraph() as s:
                s.attr(rank='same')                           # creating rank for group of block or grouping the node based cycle blocks
                for j in range(i, len(func)):
                    node_id = f'N{j}'
                    if i==len(func)-1:
                        s.node(node_id, f"\t{func[j]}\t\t\n{parameter[j]}\t\t\n{exit_condition[i]}", shape="box")  # based content shape is decided
                        break
                    elif "cycle-end" in func[j].lower():
                        s.node(node_id, f"\tCOUNT {str(inc1)} = COUNT {str(inc1)} +1\t\t\n{func[j]}\t", shape="ellipse")         #based content shape is decided
                        break
                    elif "set" in func[j].lower():
                        s.node(node_id, f"\t{func[j]}\t\t\n{parameter[j]}", shape="box")
                    elif "discharge" in func[j].lower():
                        s.node(node_id, f"\t{func[j]}\t\t\n{parameter[j]}", shape="diamond")
                    elif "charge" in func[j].lower():
                        s.node(node_id, f"\t{func[j]}\t\t\n{parameter[j]}", shape="diamond")
                    elif "rest" in func[j].lower():
                        s.node(node_id, f"\t{func[j]}\t\t\n{parameter[j]}", shape="box")
                    else:
                        s.node(node_id, f"\t{func[j]}\t\t\n{parameter[j]}", shape="ellipse")
            inc1 += 1
        cycle_head =[]
        inc=0
        for i in range(1,len(func)-1):

            if "cycle-end" in func[i].lower():
                s = parameter[i]
                l = s.split("=")
                l[0] = l[0] + " " + str(inc)
                sj = "".join([l[0], " == ", l[1]])
                exit_condition[i] = sj
                dot.edge(f'N{i}', f'N{i + 1}', taillabel=f" {exit_condition[i]}\t",minlen='3.5')                # This is for connecting that node with condition

            else:
                if exit_condition[i]:
                    dot.edge(f'N{i}', f'N{i + 1}', taillabel=f" {exit_condition[i]}\t", minlen='3.5')
                else:
                    dot.edge(f'N{i}', f'N{i + 1}', minlen='3.5')

            if "charge" in func[i].lower():
                con = str(exit_condition[i])
                print(con)
                con=con.split("\n")
                count_flag1=0
                final_con=""
                for ij in con:
                    sj = ""
                    if ">" in ij:                                                                                              # formatting string for non-passing/false condition
                        l = ij.split(">")
                        sj = "".join([l[0], " < ", l[1]])
                    elif "<" in ij:
                        l = ij.split("<")
                        sj = "".join([l[0], " > ", l[1]])
                    elif "=" in ij:
                        l = ij.split("=")
                        sj = "".join([l[0], " != ", l[1]])
                    sj+="\n"
                    final_con+=sj

                dot.edge(f'N{i}', f'N{i}', headlabel=f"\t {final_con}\t\t",minlen='3.5',tailport='w', headport='e')   #labeldistance='8',labelangle="45"

            if "cycle-start" in func[i].lower():
                cycle_head.append(f'N{i}')                     # storing head value to connect later

            if "cycle-end" in func[i].lower():
                sj=""
                if "=" in parameter[i]:                                                                                     # formatting string for non-passing/false condition
                    s = parameter[i]
                    l = s.split("=")
                    l[0] = l[0] + " " + str(inc)
                    sj = "".join([l[0], " < ", l[1]],)
                    print(sj)
                head=cycle_head[len(cycle_head)-1]
                inc+=1
                dot.edge(f'N{i}',f"{head}",headlabel=f" {sj}",labeldistance='5.0', labelangle='30',constraint='false')
                cycle_head.remove(head)

        # Checking for any left out loop and completing it
        if cycle_head:
            sj = ""
            ln=len(parameter)-1
            if "=" in parameter[ln]:  # formatting string for non-passing/false condition
                s = parameter[ln]
                l = s.split("=")
                l[0]=l[0]+" "+str(inc)
                sj = "".join([l[0], " < ", l[1], ""], )
                print(sj)
            dot.edge(f'N{len(func)-1}', f"{cycle_head[len(cycle_head)-1]}", xlabeldistance='20', xlabelangle="30", xlabel=f" {sj}",
                    constraint='false')
        dot.format = 'pdf'
        dot.render(f'{flowchart_output}\\flowchart_{output_folder}', view=True)

            # func = []
            # parameter = []
            # exit_condition = []
            # print(func, "\n", len(func))
            # for i in table:
            #     func.append(i[1])
            #     parameter.append(i[2])
            #     exit_condition.append(i[3])
            # dot = Digraph()
            # dot.attr(rankdir='LR')                                # LR: Left to Right
            # dot.attr(rank='same')                               # dot.attr(size='5,20')                               # Size in inches (Width x Height)
            # # dot.attr(ratio='compress')                            # Try to fit the content in the size
            # dot.attr(splines='ortho')                          # Straighter lines
            # dot.attr(nodesep='0.4')  # default is around 0.25
            # dot.attr(ranksep='1.0')  # default is around 0.5
            # dot.attr(concentrate='false')
            # for i in range(0, len(func)):                         # creating the nodes or block one by one
            #     with dot.subgraph() as s:
            #         s.attr(rank='same')                           # creating rank for group of block or grouping the node based cycle blocks
            #         for j in range(i, len(func)):
            #             node_id = f'N{j}'
            #             if i==len(func)-1:
            #                 s.node(node_id, f"{func[j]},\n{parameter[j]}", shape="box")  # based content shape is decided
            #                 break
            #             if "cycle-end" in func[j].lower():
            #                 s.node(node_id, f"{func[j]},\n{parameter[j]}", shape="box")         #based content shape is decided
            #                 break
            #             elif "set" in func[j].lower():
            #                 s.node(node_id, f"{func[j]},\n{parameter[j]}", shape="box")
            #             elif "discharge" in func[j].lower():
            #                 s.node(node_id, f"{func[j]},\n{parameter[j]}", shape="diamond")
            #             elif "charge" in func[j].lower():
            #                 s.node(node_id, f"{func[j]},\n{parameter[j]}", shape="diamond")
            #             elif "rest" in func[j].lower():
            #                 s.node(node_id, f"{func[j]},\n{parameter[j]}", shape="box")
            #             else:
            #                 s.node(node_id, f"{func[j]},\n{parameter[j]}", shape="ellipse")

            # cycle_head =[]

            # for i in range(len(func)-1):

            #     dot.edge(f'N{i}', f'N{i + 1}', headlabel=f"{exit_condition[i]}",minlen='3.5',labelangle='45')                 # This is for connecting that node with condition

            #     if "charge" in func[i].lower():
            #         con = str(exit_condition[i])
            #         sj = ""
            #         if ">" in con:                                                                                              # formatting string for non-passing/false condition
            #             l = con.split(">")
            #             sj = "".join([l[0], " < ", l[1]])
            #         elif "<" in con:
            #             l = con.split("<")
            #             sj = "".join([l[0], " > ", l[1]])
            #         elif "=" in con:
            #             l = con.split("=")
            #             sj = "".join([l[0], " != ", l[1]],)
            #         dot.edge(f'N{i}', f'N{i}', headlabel=f" {sj}",xlabeldistance='5',tailport='w', headport='e',
            #                 constraint='false')

            #     if "cycle-start" in func[i].lower():
            #         cycle_head.append(f'N{i}')                     # storing head value to connect later

            #     if "cycle-end" in func[i].lower():
            #         sj=""
            #         if "=" in parameter[i]:                                                                                     # formatting string for non-passing/false condition
            #             s = parameter[i]
            #             l = s.split("=")
            #             sj = "".join([l[0], " != ", l[1],"           ."],)
            #             print(sj)
            #         head=cycle_head[len(cycle_head)-1]
            #         dot.edge(f'N{i}',f"{head}", xlabeldistance='20',xlabelangle="30",xlabel=f"{sj}",constraint='false')
            #         cycle_head.remove(head)
            # if cycle_head:
            #     # with dot.subgraph() as s:
            #     #     s.attr(rank='same')
            #     #     s.node(f'N{len(func)-1}')
            #     #     s.node("", shape="plaintext", width="0", height="0")
            #     #     s.node(f"{cycle_head[len(cycle_head)-1]}")
            #     # # Make an invisible edge A -> L1 -> B
            #     # dot.edge(f'N{len(func)-1}', "L1",  style="invis")
            #     # dot.edge("L1", f"{cycle_head[len(cycle_head)-1]}", style="invis" )
            #     sj = ""
            #     ln=len(parameter)-1
            #     if "=" in parameter[ln]:  # formatting string for non-passing/false condition
            #         s = parameter[ln]
            #         l = s.split("=")
            #         sj = "".join([l[0], " != ", l[1], ""], )
            #         print(sj)
            #     dot.edge(f'N{len(func)-1}', f"{cycle_head[len(cycle_head)-1]}", xlabeldistance='20', xlabelangle="30", xlabel=f"{sj}",
            #             constraint='false')
            # dot.format = 'pdf'
            # dot.render(f'{flowchart_output}\\flowchart_{output_folder}', view=True)