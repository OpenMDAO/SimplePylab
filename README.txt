To run:

cscript pylab.vbs

The script will:
    - Download and install python
    - Download pylab_virtualenv.py
    - Create a virtual environment using pylab_virtualenv.py
    - Install the following packages into the virutal environment:
        - Numpy
        - Scipy
        - Matplotlib
        - Ipython

Once installed, to activate pylab:

pylab\Scripts\activate

To deactivate the environment just type deactivate, followed by return.

If you have versions of Python and the pylab installer on your machine that you would like to use, rather than downloading them, just add their respective paths to the PATH environment variable.

To add Python to your path:

set PATH=%PATH%;path-to-directory-with-python.exe

Similarly, you can do the same for pylab_virtualenv.py:

set PATH=%PATH%;path-to-directory-with-pylab_virtualenv.py
