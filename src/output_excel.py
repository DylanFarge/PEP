import pandas as pd
import openpyxl as xl

class ExcelOutput:
    def __init__(self, groups, self_rate, lookup):
        self.lookup = lookup        
        self.self_rate = self_rate
        self.df = pd.DataFrame(columns=["Ratings","Group"], index=groups["ID number"].tolist())

    def add_rating(self, rator, ratee, symbol):
        if rator == "22681167":
            ...
        rating = self.df.loc[rator]
        if type(rating["Ratings"]) == dict:
            self.df.loc[rator]["Ratings"][str(ratee)] = symbol
        else:
            self.df.loc[rator]["Ratings"] = {str(ratee): symbol}
                                
    def _generate_excel_workbook(self, groups):

        def lookup_func(cell):
            string = "="
            for key, value in self.lookup.items():
                string += f'IF({cell} = "{key}", {value},'
            string += "0"
            for i in range(len(self.lookup)):
                string += ")"
            return string
    
        self.wb = xl.Workbook()

        for group in groups.groupby("Group"):
            st = self.wb.create_sheet(group[0])
            members = sorted(group[1]["ID number"].astype(int).tolist())
            # group_size = len(members)
            
            for i, m in enumerate(members):
                st.cell(row=i+3, column=1, value=m)
                st.cell(row=2, column=(i*2)+2, value=m).alignment = xl.styles.Alignment(horizontal="center")
                st.merge_cells(start_row=2, start_column=(i*2)+2, end_row=2, end_column=(i*2)+3)

                m = str(m)
                if group[1][group[1]["ID number"] == m]["Flag"].values[0] == True:
                    print(f"Flagged {m}")
                    self.add_rating(m, m, [k for k,v in self.lookup.items() if v == 0][0])
                    for other in members:
                        if str(other) != m:
                            self.add_rating(m, other, [k for k,v in self.lookup.items() if v == 100][0])

                    # highlight whole row if flagged
                    for j in range(len(members)*3):
                        st.cell(row=i+3, column=j+1).fill = xl.styles.PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                for j, n in enumerate(members):
                    n = str(n)
                    if m == n and not self.self_rate:
                        # Do something where we don't rate ourselves
                        st.cell(row=i+3, column=(j*2)+2, value="-").alignment = xl.styles.Alignment(horizontal="right")
                        st.cell(row=i+3, column=(j*2)+3, value="-").alignment = xl.styles.Alignment(horizontal="right")
                    
                    else:
                        cell = xl.utils.get_column_letter((j*2)+2) + str(i+3)
                        st.cell(row=i+3, column=(j*2)+2, value=self.df.loc[m]["Ratings"][n]).alignment = xl.styles.Alignment(horizontal="right")
                        st.cell(row=i+3, column=(j*2)+3, value=lookup_func(cell))

        # remove default sheet
        self.wb.remove(self.wb["Sheet"])
    
    def save(self, groups, dir):
        self._generate_excel_workbook(groups)
        self.wb.save(dir)