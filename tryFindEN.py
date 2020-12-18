id1 = "EN"
myString = "http://142.104.193.65:8080/browse/EN-52575"

myID = myString[myString.find(id1) : ]
print(myID)