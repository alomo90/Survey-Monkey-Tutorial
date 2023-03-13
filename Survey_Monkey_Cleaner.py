# %%
import pandas as pd
import os

# %%
pwd = os.getcwd()

# %%
dataset = pd.read_excel(
    pwd + "/2023 DEI Survey - Edited.xlsx", sheet_name="Edited_Data"
)
dataset

# %%
dataset_modified = dataset.copy()
dataset_modified

# %%
dataset_modified.columns

# %%
columns_to_drop = [
    "Collector ID",
    "Start Date",
    "End Date",
    "IP Address",
    "Email Address",
    "First Name",
    "Last Name",
    "IREM ID",
    "ZIP Code",
    "Country",
]
columns_to_drop

# %%
dataset_modified = dataset_modified.drop(columns=columns_to_drop)
dataset_modified

# %%
id_vars = list(dataset_modified.columns)[0:4]
value_vars = list(dataset_modified.columns)[4:]
# value_vars

# %%
dataset_melted = dataset_modified.melt(
    id_vars=id_vars,
    value_vars=value_vars,
    var_name="Question + Subquestion",
    value_name="Answer",
)
# dataset_melted

# %%
questions_import = pd.read_excel(
    pwd + "/2023 DEI Survey - Edited.xlsx", sheet_name="Question"
)
questions_import

# %%
questions = questions_import.copy()
questions.drop(columns=["Raw Question", "Raw Subquestion", "Subquestion"], inplace=True)

# %%
questions

# %%
questions.dropna(inplace=True)

# %%
questions

# %%
dataset_merged = pd.merge(
    left=dataset_melted,
    right=questions,
    how="left",
    left_on="Question + Subquestion",
    right_on="Question + Subquestion",
)
# print(dataset_merged.shape)
print("Original Data", len(dataset_melted))
print("Merged Data", len(dataset_merged))
dataset_merged

# %%
respondents = dataset_merged[dataset_merged["Answer"].notna()]
respondents = respondents.groupby("Question")["Respondent ID"].nunique().reset_index()
respondents.rename(columns={"Respondent ID": "Respondents"}, inplace=True)
respondents

# %%
dataset_merged_two = pd.merge(
    left=dataset_merged,
    right=respondents,
    how="left",
    left_on="Question",
    right_on="Question",
)
# print(dataset_merged.shape)
print("Original Data", len(dataset_melted))
print("Merged Data", len(dataset_merged))
dataset_merged_two

# %%
same_answer = dataset_merged  # [dataset_merged["Answer"].notna()]
same_answer = (
    same_answer.groupby(["Question + Subquestion", "Answer"])["Respondent ID"]
    .nunique()
    .reset_index()
)
same_answer.rename(columns={"Respondent ID": "Same Answer"}, inplace=True)
same_answer

# %%
dataset_merged_three = pd.merge(
    left=dataset_merged_two,
    right=same_answer,
    how="left",
    left_on=["Question + Subquestion", "Answer"],
    right_on=["Question + Subquestion", "Answer"],
)
dataset_merged_three["Same Answer"].fillna(0, inplace=True)
# print(dataset_merged.shape)
print("Original Data", len(dataset_merged_two))
print("Merged Data", len(dataset_merged_three))
dataset_merged_three

# %%
output = dataset_merged_three.copy()


# %%
output.to_excel(pwd + "/2023 DEI Survey - Clean.xlsx", index=False)
