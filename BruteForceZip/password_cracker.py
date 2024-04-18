import itertools
import rarfile

def extract_file(rar_file, password):
    try:
        rar_file.extractall(pwd=password)
        return True
    except:
        return False

def main():
    rar_file = rarfile.RarFile('K:\PythonScripts\BruteForceZip\Vanguard-317-master.rar')
    characters = 'abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*()'
    min_length = 1
    max_length = 10
    attempt_counter = 0

    for length in range(min_length, max_length + 1):
        for guess in itertools.product(characters, repeat=length):
            password = ''.join(guess)
            attempt_counter += 1
            if extract_file(rar_file, password.encode()):
                print(f'Password found: {password} after {attempt_counter} attempts.')
                return

    print('Password not found')

if __name__ == '__main__':
    main()
