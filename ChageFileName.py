import os
import fnmatch

def main():
	pathiter = (os.path.join(root, filename)
	    for root, _, filenames in os.walk('.')
	    for filename in filenames
	)
	for path in pathiter:
	    newname =  path.replace('+', ' ')
	    if newname != path:
	        os.rename(path,newname)

if __name__ == '__main__':
	main()