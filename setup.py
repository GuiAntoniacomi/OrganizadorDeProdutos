from setuptools import setup

setup(
    name='your_app_name',  # Nome da sua aplicação
    version='1.0',          # Versão da sua aplicação
    description='Descrição da sua aplicação',
    author='Seu Nome',
    packages=['your_package_name'], # Se você tiver um pacote, adicione o nome aqui
    data_files=[('FrontEnd', ['FrontEnd/img0.png', 'FrontEnd/img1.png', 'FrontEnd/img2.png', 'FrontEnd/img3.png', 'FrontEnd/img_textBox0.png', 'FrontEnd/img_textBox1.png', 'FrontEnd/img_textBox2.png', 'FrontEnd/background.png'])],
    options={
        'build_exe': {
            'include_files': [
                ('FrontEnd', 'FrontEnd')
            ]
        }
    },
    pyinstaller_options={
        'distutils_args': [
            '--onefile',
            '--windowed',
        ],
        'hiddenimports': [
            'PIL.Image',
            'PIL.ImageTk'
        ],
        'datas': [
            ('FrontEnd/img0.png', 'FrontEnd'),
            ('FrontEnd/img1.png', 'FrontEnd'),
            ('FrontEnd/img2.png', 'FrontEnd'),
            ('FrontEnd/img3.png', 'FrontEnd'),
            ('FrontEnd/img_textBox0.png', 'FrontEnd'),
            ('FrontEnd/img_textBox1.png', 'FrontEnd'),
            ('FrontEnd/img_textBox2.png', 'FrontEnd'),
            ('FrontEnd/background.png', 'FrontEnd'),
        ]
    },
)