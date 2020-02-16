# mt_detect
A script to compare documents using Difflib.

Wiki needs to be updated.

Compilation notes (these steps necessary to ensure that package is small; otherwise, 200+ MB size)
1. Created virtual environment (virtualenv doc_compare). Deactivate base anaconda environment
2. In virtual env, installed all libraries (xlrd, python-pptx, pypiwin32)
3. In virtual env, installed pyinstaller
4. Using pyinstaller -w -F (meaning not windowed and onefile), compiled script  
  
Usage notes  
1. Keep model files in folder named "model"  
2. Keep solution files in folders named "test_{solution}", e.g. abc would be "test_abc"  
3. All solution files should have same names as model files  
4. Keep doc_compare script file in the same folder as model and solution folders and double click to run  
