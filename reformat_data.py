import os


def titlecase_mahasiswa(file='data/Daftar_Mahasiswa.txt'):
    with open(file, 'r') as f_in:
        lines = f_in.readlines()
    
    lines = [line.title() for line in lines]
    filename, ext = os.path.splitext(file)
    with open(f'{filename}_Converted{ext}', 'w') as f_out:
        f_out.writelines(lines)

def generate_email_from_name(file='data/Daftar_Dosen.txt'):
    with open(file, 'r') as f_in:
        lines = f_in.readlines()

    filename, ext = os.path.splitext(file)
    lines = ['_'.join(line.lower().split()) + '@gmail.com\n' for line in lines]
    with open(f'{filename}_Email{ext}', 'w') as f_out:
        f_out.writelines(lines)

def uppercase_topic(file='data/Daftar_Topik.txt'):
    with open(file, 'r') as f_in:
        lines = f_in.readlines()
    
    lines = [line.upper() for line in lines]
    filename, ext = os.path.splitext(file)
    with open(f'{filename}_Converted{ext}', 'w') as f_out:
        f_out.writelines(lines)

if __name__ == '__main__':
    titlecase_mahasiswa()
    generate_email_from_name()
    uppercase_topic()
