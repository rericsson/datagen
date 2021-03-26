from setuptools import setup

setup(
        name='datagen',
        version='0.1',
        py_modules=['datagen'],
        install_requires=[
                'click==7.1.1',
                'PyYAML==5.4',
                'XlsxWriter==1.2.8'
            ],
        entry_points='''
            [console_scripts]
            datagen=datagen:datagen
        '''
        )
