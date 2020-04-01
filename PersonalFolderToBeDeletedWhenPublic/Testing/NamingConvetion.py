import os
clear = lambda: os.system('cls') #on Windows System
clear()

def GenerateVersion():

    LatesIssues = input("Please Enter The Number Of Open Issues: ")
    NoOfCommits = input("Please Enter The Number Of Commits: ")
    NoOfReleases = input("Please Enter The Number Of Releases: ")
    Version = "C-19."+LatesIssues+"."+NoOfCommits+"."+NoOfReleases
    print("Version is:", Version)
    return Version

UserOkay = False
while UserOkay == False:
    Version = GenerateVersion()
    tmpinput = input("Is This Version Name Okay?: Y/N")
    if tmpinput == "Y" or tmpinput == "y":
        UserOkay = True
        clear()
        print("Thank You For Using Owen's Version Generator")
        print("Your Version Is:",Version)
    else:
        print("Regenerating")


