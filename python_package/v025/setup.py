import setuptools    

setuptools.setup(
    name='excel_to_dataframe',
    version='0.2.5',
    author='Nelson Rossi Bittencourt',
    author_email='nbittencourt@hotmail.com',
    description='Excel to Pandas or Microsoft Dataframe',
    long_description='C++ dll to converts Excel sheets to Pandas or Microsoft dataframes',
    long_description_content_type="text/markdown",
    url='https://github.com/nbittencourt/excel_to_dataframe',
    license='MIT',
    packages=['excel_to_dataframe'],
	include_package_data=True,
	package_data={'':['excel_to_df.dll'],'benchmarks\\':['benchmarks\\readme.txt']},
    install_requires=['pandas'],
)