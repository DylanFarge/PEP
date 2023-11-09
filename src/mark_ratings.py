import os, sys
import pandas as pd
import numpy as np
from src.output_excel import ExcelOutput
# from output_excel import ExcelOutput

open_log_file : bool = True
outdir : str = "."

def out(message : str, type="string") :
    '''
    This method prints a message to the console and writes it to the log file.
    
    Parameters
    ----------
    message : str
        The message to print and write to the log file.
        
    type : str (optional)
        The type of message to print. Either "string" or "dataframe". Default is "string".
    
    Returns
    -------
    None
    '''
    global open_log_file, outdir
    if open_log_file:
        os.mkdir(outdir) if not os.path.exists(outdir) else None
        write = 'w'
        open_log_file = False
    else:
        write = 'a'
    with open(f"{outdir}/log.txt", write) as f:
        if type == "string":
            f.write(message+"\n")
            # print(message)
        else:
            f.write(message.to_string()+"\n")
            # print(message.to_string())

def get_data(path : str) -> pd.DataFrame:
    '''
    This method reads the data file and returns a dataframe with the data.
    
    Parameters
    ----------
    path : str
        The path to the data file.
        
    Returns
    -------
    A dataframe with the data.
    '''
    try:
        data = pd.read_csv(path)
    except:
        raise IOError("Could not find file: "+path+"\n Please make sure that the file path exists and correct.")
    return data


def get_groups(path : str) -> pd.DataFrame:
    '''
    This method reads the groups file and returns a dataframe with the groups and their members.
    pending : list
    
    Parameters
    ----------
    path : str
        The path to the groups file.
        
    Returns
    -------
    A dataframe with the groups and their members.
    '''
    try:
        groups = pd.read_excel(path)
    except Exception as e:
        print(e)
        raise IOError("Could not find file: "+path+"\n Please make sure that the file path exists and correct.")
    groups['Size'] = groups.groupby('Group')['Group'].transform('count')
    # make sure that the ids are all strings and not given as integers
    groups["ID number"] = list(map(str,groups["ID number"].to_list()))
    groups['Flag'] = False
    return groups

    
def try_to_correct(member : str, group_members : list) -> str:
    '''
    This method attempts to correct a student's id number if it is incorrect.
    
    Parameters
    ----------
    member : str
        The incorrect id number in which to correct.     
           
    group_members : str
        A list of all the id numbers in the same group.
        
    Returns
    -------
    The corrected id number as a string.
    '''
    if member == "-" or member == "" or member == "nan":
        out("...Could not find correct id number")
        out("****************************************************")
        return None
    
    out("Attempting to correct...")
    member = int(member)
    correction = 0
    correction_calc = 99999999
    
    for student_id in group_members:
        id_num = int(student_id)
        calc = int(str(abs(id_num - member)).rstrip('0'))
        if calc < correction_calc:
            correction = id_num
            correction_calc = calc
    
    if correction == 0:
        out("...Could not find correct id number")
        out("****************************************************")
        return None
    
    out("Changed from:\t"+str(member))
    out("Changed to: \t"+str(correction))
    out("****************************************************")
    return str(correction)
    
    
def process_students(groups : pd.DataFrame, data : pd.DataFrame, lookup : dict, settings : dict) -> list[pd.DataFrame, pd.DataFrame]:
    '''
    This method processes every student. This is where the actual calculations are done.
    
    Parameters
    ----------
    groups : pd.DataFrame
        Dataframe that contains info on all the groups
        
    data : pd.DataFrame
        The meta data that is used to calculate the ratings
        
    lookup : dict
        The lookup table to convert the ratings to numbers
        
    settings : dict
        The settings used to calculate the ratings
        
    Returns
    -------
    Both the student mean and the resulting group dataframe
    '''
    global excel
    correct_ids = settings["correct_ids"]
    student_mean = pd.DataFrame(np.zeros((1, len(groups))), columns=groups["ID number"].to_list())
    for entry in data.iterrows():
        student = str(entry[1]["ID number"])
        if student == "25081551":
            ...
        group = groups[groups["ID number"] == student]
        group_size = int(group["Size"].values[0])
        group_size = group_size if settings["self_rate"] else group_size - 1
        group = group["Group"].values[0]
        
        out("---------- Processing "+student+" "+entry[1]["Surname"]+" ----------")
        i = 0
        pending_scores = []
        pending_members = []
        duplicate = False
        j = 0 # useful if self_rate is False
        while i < group_size:
            
            try:
                    
                member = str(entry[1]["Response "+str(3*(i+j)+1)]).split(".")[0]
                rating = str(entry[1]["Response " + str(3*(i+j)+2)])
                member_group = groups[groups["ID number"] == member]["Group"]
                
                if member == student and not settings["self_rate"] and member not in pending_members:
                    pending_members.append(member)
                    out("\t"+ member+ " -> IGNORING SELF RATING")
                    j += 1
                    continue

                if member_group.empty:
                    raise ValueError("Could not find member: "+member+" in provided groups file.")
                elif member_group.values[0] != group:
                    raise ValueError("Member: "+member+" is not in the same group as student: "+student)
                elif member in pending_members:
                    duplicate = True
                    raise ValueError("Member: "+member+" has been rated twice by student: "+student)
                

                student_mean[member] += lookup[rating] / group_size
                pending_scores.append(lookup[rating] / group_size)
                pending_members.append(member)
                
                out("\t"+ member+ " -> "+ rating)
                excel.add_rating(student, member, rating)
                i += 1
                
            except Exception as e:
                print(e)
                out("********** ERROR IN ENTRY OF "+entry[1]["First name"]+" **********")
                out("Mistake detected in following input:")
                out("Member ID : "+ member)
                out("Rating    : "+ rating)
                if duplicate:
                    out("***** Duplicate rating *****")
                out("****************************************************")
                
                if rating == "nan" or rating == "" or rating == " " or rating == "-":
                    groups.loc[groups["ID number"] == student, "Flag"] = True
                    out("Student FLAGGED")
                    for already in range(len(pending_scores)):
                        student_mean[pending_members[already]] -= pending_scores[already]
                    i = group_size

                elif correct_ids and not duplicate:
                    correction = try_to_correct(member, groups[groups["Group"] == group]["ID number"].tolist())
                    
                    if correction is None:
                        groups.loc[groups["ID number"] == student, "Flag"] = True
                        out("Student FLAGGED")
                        
                        for already in range(len(pending_scores)):
                            student_mean[pending_members[already]] -= pending_scores[already]
                        i = group_size
                   
                    else:
                        entry[1]["Response "+str(3*(i+j)+1)] = correction
                        
                else:
                    groups.loc[groups["ID number"] == student, "Flag"] = True
                    out("Student FLAGGED")
                    
                    for already in range(len(pending_scores)):
                        student_mean[pending_members[already]] -= pending_scores[already]
                    i = group_size
                    
    out("--------------------------------------------")
    return student_mean, groups


def compile_results(groups : pd.DataFrame, factors : list, create : bool):
    '''
    This method is used to compile the results and output them into one exel file for the user.
    
    Parameters
    ----------
    groups : pd.DataFrame
        The dataframe in which the results are hidden within.
        
    factors : list
        The scalling factors 
        
    create : bool   
        Will create a excel file if True
        
    Returns
    -------
    None
    '''
    global outdir, excel
    excel.save(groups, outdir+"/results.xlsx")
    results = pd.DataFrame({
        "First name": groups["First name"].tolist(),
        "Surname": groups["Surname"].tolist(),
        "ID number": groups["ID number"].tolist(),
        "Group": groups["Group"].tolist(),    
        "Scaling factor": factors,
        "Flagged": ["" if i == False else "FLAGGED" for i in groups["Flag"].tolist()],
    })
    results = results[ results["Scaling factor"] != -1]    
    results.sort_values(by=["ID number"], inplace=True, ignore_index=True)
    
    out("Final scores:")
    out("--------------------------------------------")
    out(results, type="df")
    out("--------------------------------------------")
    out("FLAGGED students:")
    out("--------------------------------------------")
    out(results[results["Flagged"] == "FLAGGED"], type="df")
    out("--------------------------------------------")
    
    if create:
        if os.path.exists(f"{outdir}/scaling_factors.csv"):
            file = pd.read_csv(f"{outdir}/scaling_factors.csv")
            cols = file.columns.tolist()
            count = 0
            for col in cols:
                if "Scaling factor" in col:
                    count += 1
            newdf = pd.DataFrame(columns=["Scaling factor "+str(count+1),"Flagged "+str(count+1)])
            newdf["Scaling factor "+str(count+1)] = results["Scaling factor"]
            newdf["Flagged "+str(count+1)] = results["Flagged"]
            df = pd.concat([file, newdf], axis=1, ignore_index=False)
        else:
            df = results
            df.rename(columns={"Scaling factor": "Scaling factor 1", "Flagged": "Flagged 1"}, inplace=True)
                
        with open(f"{outdir}/scaling_factors.csv", "w") as f:
            f.write(df.to_csv(index=False))

    
def mark(args: str(list), self_rate: bool = True,settings: dict = {
    "capped_mark": 1.1,
    "correct_ids": True,
    "create_file": True,
}, lookup:dict = {
    'E': 100,
    'V': 87.5,
    'S': 75,
    'O': 62.5,
    'M': 50,
    'D': 37.5,
    'U': 25,
    'F': 12.5,
    'N': 0,
}):
    '''
    This script is designed to take in a specific format of a quiz pulled from SUNLearn, and automatically mark it.
    
    Parameters
    ----------
    args : str(list)
        str1 : The path to the group file
        str2 : The path to the peer review meta-data intended to be marked.
        str3 : The path to the output directory

    self_rate(optional) : boolean
        If True, students can rate themselves. Default is True.
        
    settings(optional) : dict
        capped_mark(float) : number used to set a cap to how high the final score can be
        correct_ids(boolean) : attempt to correct ids if possible, otherwise just flag students that have errors in input.
        create_file(boolean) : creates file with scores, but does not effect the output of the log file.
        
    lookup(optional) : dict
        This is a dictionary that will be used to compare the students' responses and marks to. The 
        default lookup table is as follows:
        'E' : 100
        'V' : 87.5,
        'S' : 75,
        'O' : 62.5,
        'M' : 50,
        'D' : 37.5,
        'U' : 25,
        'F' : 12.5,
        'N' : 0,
        
    Returns
    ----------
    None
    '''
    global outdir, excel
    outdir = args[2]
    groups = get_groups(args[0])
    data = get_data(args[1])  
    excel = ExcelOutput(groups, self_rate, lookup, settings["capped_mark"])
    settings["self_rate"] = self_rate
    student_mean, groups = process_students(groups, data, lookup, settings)
    
    # Flag students that did not complete quiz while belonging to a group
    for i in groups["ID number"].tolist():
        if int(i) not in data['ID number'].tolist() and not groups["Group"][groups["ID number"] == i].isna().values[0] :
            groups.loc[groups["ID number"] == i, "Flag"] = True
    
    # Assign 0 to flagged students and 100 to members in the same group
    flagged_students = groups[groups["Flag"] == True]
    for entry in flagged_students.iterrows():
        group = entry[1]["Group"]
        group_members = groups[groups["Group"] == group]["ID number"].tolist()
        for i in group_members:
            if i == entry[1]["ID number"]:
                student_mean[i] += 0
            else:
                student_mean[i] += (100 / len(group_members))

    # Calculate team means
    group_names = groups["Group"].unique().tolist()
    team_mean = pd.DataFrame(np.zeros((1, len(group_names))), columns=group_names)
    for g in group_names:
        group = groups[groups["Group"] == g]
        if group.empty:
            continue
        size = int(group["Size"].values[0])
        team_mean[g] = student_mean[group["ID number"].tolist()].sum(axis=1) / size
    
    # Calculate scaling factor
    factors = []
    for entry in groups.iterrows():
        group = entry[1]["Group"]
        if team_mean[group].values[0] == 0:
            factors.append(-1)
            continue
        factor = np.minimum(student_mean[entry[1]["ID number"]].values[0] / team_mean[group].values[0], settings["capped_mark"])
        factors.append(factor.round(2))        
        
    # Compiles and prints results
    compile_results(groups, factors, settings["create_file"])
    

if __name__ == "__main__":
    if len(sys.argv) > 1:
        mark(sys.argv[1:], self_rate=True)
    else:
        paths = [
            "files/groups.xlsx",
            "files/p1.csv",
            "output"
        ]
        mark(paths, self_rate=False)