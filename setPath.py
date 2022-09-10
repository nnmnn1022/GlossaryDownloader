def run(path):
    del_list = ['"', "'", '& ']
    for char in del_list :
        path = path.replace(char, '')
    path = str(path.replace('\\', '/'))

    return path

if __name__ == '__main__' :
    run()