from setuptools import setup

setup(
        name='datagen',
        version='0.1',
        py_modules=['datagen'],
        install_requires=[
                'attrs==19.3.0',
                'click==7.1.1',
                'et-xmlfile==1.0.1',
                'Faker==4.0.3',
                'jdcal==1.4.1',
                'more-itertools==8.2.0',
                'openpyxl==3.0.3',
                'packaging==20.3',
                'pluggy==0.13.1',
                'py==1.8.1',
                'pyparsing==2.4.7',
                'pytest==5.4.1',
                'python-dateutil==2.8.1',
                'PyYAML==5.3.1',
                'six==1.14.0',
                'text-unidecode==1.3',
                'wcwidth==0.1.9',
                'XlsxWriter==1.2.8'
            ],
        entry_points='''
            [console_scripts]
            datagen=datagen:datagen
        '''
        )