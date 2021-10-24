from time import sleep
import OpenOPC

while True:
    opc = OpenOPC.client()
    opc.connect('Graybox.Simulator.1')
    valor = opc['numeric.sin.uint8']
    opc.close()

    file = open("opcvalue.txt","w")
    file.write(str(valor))
    file.close()
    sleep(1)

