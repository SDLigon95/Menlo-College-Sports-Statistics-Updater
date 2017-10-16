from distutils.core import setup
import py2exe

setup_dict = dict(
        console=['SportsGrabber.py'],
        windows=[{
                "script": "SportsGrabber.py",
                "icon_resources": [(1, "SportsGrabber.ico")],
            }],
        options={
        'py2exe': 
        {
            'includes': ['lxml.etree', 'lxml._elementpath']
        }
        }

)
setup(**setup_dict)
setup(**setup_dict)