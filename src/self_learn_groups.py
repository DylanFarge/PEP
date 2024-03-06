import pandas as pd
import numpy as np

def learn_groups(path, to_file):
    # Load the data
    df_raw = pd.read_csv(path)

    # Get the member ids and the responses
    responses_col = [col for col in df_raw.columns if "Response" in col] 
    ids = [responses_col[i] for i in range(0, len(responses_col), 3)]
    ids.insert(0, "ID number")
    df_ids = df_raw[ids]

    # Removing NaNs and '-' from the data, and self-rating duplicates
    ids = df_ids.to_numpy(na_value='-', dtype=str).tolist()
    for i in range(len(ids)):
        ids[i] = set([int(float(x)) for x in ids[i] if x != '-'])
        # unique_members = list(set([int(x) for x in ids[i] if x != '-']))
        # ids[i] = sorted(unique_members)

    # Associate the members to groups
    found_group = True
    while found_group:
        found_group = False

        for i in range(len(ids)):
            for j in range(i+1, len(ids)):
                if len(ids[i].intersection(ids[j])) > 0:
                    found_group = True
                    ids[i] = ids[i].union(ids[j])
                    ids[j] = set()
                    break
            if found_group:
                break
    ids = [group for group in ids if len(group) > 0]

    # Check if member id exists as ID number
    legit_ids = set(df_raw["ID number"].to_numpy())
    print("Removed members that don't exist in the ID number column.-------------")
    for group in ids:
        remove = []
        for member in group:
            if member not in legit_ids:
                print(f"{member}")
                remove.append(member)    
        group.difference_update(remove)
    print("----------------------------------")

    # Check if there are any duplicate ids
    seen = set()
    dup = set()

    for group in ids:
        for member in group:
            if member in seen:
                dup.add(member)
            else:
                seen.add(member)
    if len(dup) > 0:
        print("IDs shared in more than one group.----------------")
        print(dup)
        raise ValueError("NEED TO FIX: IDs shared in more than one group.")


    # Count the group sizes
    sizes = {}
    for group in ids:
        size = len(group)
        if size in sizes:
            sizes[size] += 1
        else:
            sizes[size] = 1
    
    # Name the groups
    sizes = dict(sorted(sizes.items(), key=lambda item: item[1], reverse=True))
    groups = {}
    group_names = 0
    for size, count in sizes.items():
        print(f"--------------Group size: {size} - Count: {count}--------------------")
        sized_groups = [group for group in ids if len(group) == size]
        for group in sized_groups:
            group_names += 1
            groups[f"CalGr{group_names:02d}"] = sorted(group)
            print(f"CalGr{group_names:02d} - {sorted(group)}")

    # Output to csv file
    if to_file:
        all_members = []
        for group in groups.values():
            all_members.extend(group)
        
        df = df_raw[df_raw["ID number"].isin(all_members)][["First name", "Last name", "ID number"]]
        df.sort_values(by=["First name", "Last name"], inplace=True, ignore_index=True) 
        group_allocation = []
        group_sizes = []

        for student in  df["ID number"].to_list():
            for group, members in groups.items():
                if student in members:
                    group_allocation.append(group)
                    group_sizes.append(len(members))
        
        df["Group"] = group_allocation
        df["Group size"] = group_sizes

        df.to_excel("output/calculated_group_allocations.xlsx", index=False)

if __name__ == '__main__':
    learn_groups("~/PEP/2024-MT00548-Project 1 Chat peer rating-responses.csv", to_file=True) #XXX