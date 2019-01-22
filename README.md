==============
WORKFLOW GIT
==============
### Create a brach, make the changes, push the changes to the branch.
* First, checkout the branch that you want to make a new branch out of.
* Create a branch
	git branch <branch name>
* switch to new branch
	git checkout <new branch name>
* Then, make all the changes.
* git status (to see the changes you have made)
* git add -A (to add everything to statge area)
* git commit -m "message"
* git push -u origin <branch name>

### Merge branch to master
* git checkout master
* git pull origin master
* git merge <branch name>
* git push origin master

### Delete the branch from local and remote
* git branch -d <branch name> (delete from local)
* git push origin --delete <branch name> (from remote)