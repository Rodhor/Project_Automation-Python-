import argparse
import os
import shutil
from dotenv import load_dotenv
import github

load_dotenv()
project_info = []

def args():
    arg = argparse.ArgumentParser()
    print('To exit the projectcreation, write "Cancel" \n')
    
    while True:
        arg.name = input("Projectname: ").lower()
        if arg.name =="":
            print("\nPlease input a Projectname!\n")
        elif arg.name =="cancel":
            cancel()
        else:
            break
    
    while True:
        arg.language = input("VBA or Python? ").lower()
        if arg.language == "vba":
            break
        elif arg.language =="python":
            break
        elif arg.language == "cancel":
            cancel()
        else:
            print("Please select either VBA or Python")
    
    while True:
        arg.github = input("Create a Github repository? (y/n): ").lower()
        if arg.github == "cancel":
            cancel()
        elif arg.github == 'n':
            break
        elif arg.github =='y':
            break
        else:
            print("Invalid response")

    return arg.name, arg.language, arg.github

def cancel():
    print('\nProjectcreation was terminated! \n')
    exit()

def VBA(name):
    folderpath = os.getenv('FOLDERPATH_VBA')
    Folderpath_file = os.getenv('FOLDERPATH_VBA_FILE')
    new_path = os.path.join(folderpath, name)
    create_folder(new_path)
    New_File_path = shutil.copy2(Folderpath_file,new_path)
    VBAProject = f'{new_path}{os.sep}{name}.xlsm'
    os.rename(New_File_path,VBAProject)
    os.startfile(VBAProject)

def PYTHON(name):
    folderpath = os.getenv('FOLDERPATH_PYTHON')
    new_path = os.path.join(folderpath, name)
    create_folder(new_path)
    file = f'{new_path}{os.sep}{name}.py'
    open(file,'x')
    os.startfile(file)

def git(remote, name):
    os.system('git init')
    os.system('git branch -M Master')
    if remote =='y':
        github(name)
    else:
        os.system('git add .')
        os.system('git commit -m "Initial Commit" ')

def github(repo_name):
    user = Github(os.getenv('GITHUB_ACCESSTOKEN')).get_user()
    user.create_repo(name=repo_name, private=True)
    os.system(f'git remote add origin https://github.com/{user.login}/{repo_name}.git')
    os.system('git add .')
    os.system('git commit -m "Initial Commit" ')
    os.system('git push origin Master')

def create_folder(new_path):
    try:
        os.mkdir(new_path)
        os.chdir(new_path)
        open('readme.md',"x")
    except FileExistsError as err:
        raise SystemExit(err)

def Create_Project():
    project_info = args()
    name = project_info[0]
    remote = project_info[2]

    if project_info[1] == "python":
        PYTHON(name)
    elif project_info[1] == "vba":
        VBA(name)

    git(remote,name)
    
    print("\nSucces!")
    

if __name__ == '__main__':
    Create_Project()