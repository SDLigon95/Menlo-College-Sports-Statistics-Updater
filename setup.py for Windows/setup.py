from distutils.core import setup
import py2exe

setup(
        console=['SportsGrabber.py'],
        options={
        'py2exe': 
        {
            'includes': ['lxml.etree', 'lxml._elementpath']
        }
        }

)
