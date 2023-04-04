import setuptools

setuptools.setup(
    name="sales-to-bill",
    version="1.0",
    author="Thomas Lunger",
    author_email="t.lunger1@fatpipeinc.com",
    description="Updates 'In Service Dates', S/N, Tracking#, and Coloring in a Bill Trigger Sheet, based on Sales Sheet Data.",
    packages=setuptools.find_packages(),
    install_requires=[
        "pandas",
        "openpyxl",
    ],
    entry_points={
        "console_scripts": [
            "sales-to-bill=sales_to_bill.main:main",
        ],
    },
)