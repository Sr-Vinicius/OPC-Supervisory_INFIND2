import OpenOPC

opc = OpenOPC.client()
opc.connect('Graybox.Simulator.1')
print(str(opc['numeric.sin.uint8']))

opc.close()