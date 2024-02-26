import os
import shutil
from zipfile import ZipFile
import pytest

current_file = os.path.abspath(__file__)
current_dir = os.path.dirname(current_file)
resource_dir_path = os.path.join(current_dir, 'resource')
resource_zip_path = os.path.join(resource_dir_path, 'resource.zip')
files_dir_path = os.path.join(current_dir, 'tmp')


@pytest.fixture(scope='function', autouse=True)
def create_archive():
    if not os.path.exists(resource_dir_path):
        os.mkdir(resource_dir_path)
    with ZipFile(resource_zip_path, 'w') as zf:
        for file in os.listdir(files_dir_path):
            add_file = os.path.join(files_dir_path, file)
            zf.write(add_file, file)
    yield
    shutil.rmtree(resource_dir_path)