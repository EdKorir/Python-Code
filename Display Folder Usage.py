print('NB: For Drive C:, enter C:\\\, otherwise, enter the path in normal format')
path_entered = input()
print('Path entered: ',path_entered)

def Folder_usage(path):
    import os
    import seaborn as sns
    import matplotlib.pyplot as plt
    from matplotlib.pyplot import figure
    import pandas as pd
    
    os.chdir(r'C:')

    subdirs = [os.path.join(path, o) for o in os.listdir(path) if os.path.isdir(os.path.join(path,o))]
    #print(subdirs)

    def size_calc(g):
        folder_size = 0
        for dirname, _, filenames in os.walk(g):
            for filename in filenames:
                fname = os.path.join(dirname,filename)
                folder_size += (os.path.getsize(fname)/(1024*1024*1024))
        return folder_size

    subdirsize = []
    for y in subdirs:
        try:
            subdirsize.append(size_calc(y))
        except OSError as e:
            print('File Not Found')
            subdirsize.append(0)
    #print(subdirsize)   

    data = pd.DataFrame(subdirsize, index=subdirs, columns=['Folder Size (GBs)'])
    data = data.sort_values(by=['Folder Size (GBs)'], ascending=False)
    data = data.head(20)
    data

    sns.set(rc={'figure.figsize':(13, 10)})
    sns.set(style='whitegrid')
    plt.xticks(rotation = 45)
    return data, sns.barplot(x=data.index,y=data['Folder Size (GBs)'], data=data)

Folder_usage(path_entered)