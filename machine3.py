from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score
import csv
import pickle

def splitin(item):
    array1 = item[:4]
    array2 = item[4:]
    caractere='['
    caractere1=']'
    array1 = [string.replace(caractere, '') for string in array1]
    array1 = [string.replace(caractere1, '') for string in array1]
    arr_int1 = (list(map(int, array1)))
    return arr_int1,array2
  

with open('d1.csv', newline='') as csvfile:
    X=[]
    Y=[]
    next(csvfile)
    data = (csv.reader(csvfile))
    dados = [[linha[:12], linha[12:]] for linha in data]
    for item in dados:
        for element in item:
            if(len(element)):
                aux,aux1=splitin(element)
                X.append(aux)
                Y.append(aux1)

model = DecisionTreeClassifier().fit(X,Y)
arr1=([0, 0, 0, 0],[1, 0, 0, 0],[1, 1, 0, 0],[1, 1, 1, 0],[1, 1, 1, 1],[1, 1, 1, 1],[0, 0, 0, 0],[1, 0, 0, 0],[1, 1, 0, 0],[1, 1, 1, 0],[1, 1, 1, 1])
outp1=(["00FF9900","00FFCC00","FFE68D","FFFF99","00FFFFFF","00FFFFFF","00FF9900","00FFCC00","FFE68D","FFFF99","00FFFFFF"])
y_pred = model.predict(arr1)
accuracy = accuracy_score(outp1, y_pred)
print(y_pred)
print("Acurácia:", accuracy)
with open('modelo.pkl', 'wb') as f:
    pickle.dump(model, f)

# Obtendo as colunas do DataFrame

# Imprimindo as colunas










