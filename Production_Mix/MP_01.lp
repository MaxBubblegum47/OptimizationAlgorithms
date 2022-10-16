0\ Model MP_01
\ LP format - for model browsing. Use MPS format to capture full model detail.
Maximize
  10 sales[January,P1] + 6 sales[January,P2] + 8 sales[January,P3]
   + 4 sales[January,P4] + 11 sales[January,P5] + 9 sales[January,P6]
   + 3 sales[January,P7] + 10 sales[February,P1] + 6 sales[February,P2]
   + 8 sales[February,P3] + 4 sales[February,P4] + 11 sales[February,P5]
   + 9 sales[February,P6] + 3 sales[February,P7] + 10 sales[March,P1]
   + 6 sales[March,P2] + 8 sales[March,P3] + 4 sales[March,P4]
   + 11 sales[March,P5] + 9 sales[March,P6] + 3 sales[March,P7]
   + 10 sales[April,P1] + 6 sales[April,P2] + 8 sales[April,P3]
   + 4 sales[April,P4] + 11 sales[April,P5] + 9 sales[April,P6]
   + 3 sales[April,P7] + 10 sales[May,P1] + 6 sales[May,P2]
   + 8 sales[May,P3] + 4 sales[May,P4] + 11 sales[May,P5] + 9 sales[May,P6]
   + 3 sales[May,P7] + 10 sales[June,P1] + 6 sales[June,P2]
   + 8 sales[June,P3] + 4 sales[June,P4] + 11 sales[June,P5]
   + 9 sales[June,P6] + 3 sales[June,P7] - 0.5 inventory[January,P1]
   - 0.5 inventory[January,P2] - 0.5 inventory[January,P3]
   - 0.5 inventory[January,P4] - 0.5 inventory[January,P5]
   - 0.5 inventory[January,P6] - 0.5 inventory[January,P7]
   - 0.5 inventory[February,P1] - 0.5 inventory[February,P2]
   - 0.5 inventory[February,P3] - 0.5 inventory[February,P4]
   - 0.5 inventory[February,P5] - 0.5 inventory[February,P6]
   - 0.5 inventory[February,P7] - 0.5 inventory[March,P1]
   - 0.5 inventory[March,P2] - 0.5 inventory[March,P3]
   - 0.5 inventory[March,P4] - 0.5 inventory[March,P5]
   - 0.5 inventory[March,P6] - 0.5 inventory[March,P7]
   - 0.5 inventory[April,P1] - 0.5 inventory[April,P2]
   - 0.5 inventory[April,P3] - 0.5 inventory[April,P4]
   - 0.5 inventory[April,P5] - 0.5 inventory[April,P6]
   - 0.5 inventory[April,P7] - 0.5 inventory[May,P1]
   - 0.5 inventory[May,P2] - 0.5 inventory[May,P3] - 0.5 inventory[May,P4]
   - 0.5 inventory[May,P5] - 0.5 inventory[May,P6] - 0.5 inventory[May,P7]
   - 0.5 inventory[June,P1] - 0.5 inventory[June,P2]
   - 0.5 inventory[June,P3] - 0.5 inventory[June,P4]
   - 0.5 inventory[June,P5] - 0.5 inventory[June,P6]
   - 0.5 inventory[June,P7]
Subject To
 R0: - x[January,P1] + sales[January,P1] + inventory[January,P1] = 0
 R1: - x[January,P2] + sales[January,P2] + inventory[January,P2] = 0
 R2: - x[January,P3] + sales[January,P3] + inventory[January,P3] = 0
 R3: - x[January,P4] + sales[January,P4] + inventory[January,P4] = 0
 R4: - x[January,P5] + sales[January,P5] + inventory[January,P5] = 0
 R5: - x[January,P6] + sales[January,P6] + inventory[January,P6] = 0
 R6: - x[January,P7] + sales[January,P7] + inventory[January,P7] = 0
 R7: - x[February,P1] + sales[February,P1] - inventory[January,P1]
   + inventory[February,P1] = 0
 R8: - x[February,P2] + sales[February,P2] - inventory[January,P2]
   + inventory[February,P2] = 0
 R9: - x[February,P3] + sales[February,P3] - inventory[January,P3]
   + inventory[February,P3] = 0
 R10: - x[February,P4] + sales[February,P4] - inventory[January,P4]
   + inventory[February,P4] = 0
 R11: - x[February,P5] + sales[February,P5] - inventory[January,P5]
   + inventory[February,P5] = 0
 R12: - x[February,P6] + sales[February,P6] - inventory[January,P6]
   + inventory[February,P6] = 0
 R13: - x[February,P7] + sales[February,P7] - inventory[January,P7]
   + inventory[February,P7] = 0
 R14: - x[March,P1] + sales[March,P1] - inventory[February,P1]
   + inventory[March,P1] = 0
 R15: - x[March,P2] + sales[March,P2] - inventory[February,P2]
   + inventory[March,P2] = 0
 R16: - x[March,P3] + sales[March,P3] - inventory[February,P3]
   + inventory[March,P3] = 0
 R17: - x[March,P4] + sales[March,P4] - inventory[February,P4]
   + inventory[March,P4] = 0
 R18: - x[March,P5] + sales[March,P5] - inventory[February,P5]
   + inventory[March,P5] = 0
 R19: - x[March,P6] + sales[March,P6] - inventory[February,P6]
   + inventory[March,P6] = 0
 R20: - x[March,P7] + sales[March,P7] - inventory[February,P7]
   + inventory[March,P7] = 0
 R21: - x[April,P1] + sales[April,P1] - inventory[March,P1]
   + inventory[April,P1] = 0
 R22: - x[April,P2] + sales[April,P2] - inventory[March,P2]
   + inventory[April,P2] = 0
 R23: - x[April,P3] + sales[April,P3] - inventory[March,P3]
   + inventory[April,P3] = 0
 R24: - x[April,P4] + sales[April,P4] - inventory[March,P4]
   + inventory[April,P4] = 0
 R25: - x[April,P5] + sales[April,P5] - inventory[March,P5]
   + inventory[April,P5] = 0
 R26: - x[April,P6] + sales[April,P6] - inventory[March,P6]
   + inventory[April,P6] = 0
 R27: - x[April,P7] + sales[April,P7] - inventory[March,P7]
   + inventory[April,P7] = 0
 R28: - x[May,P1] + sales[May,P1] - inventory[April,P1] + inventory[May,P1]
   = 0
 R29: - x[May,P2] + sales[May,P2] - inventory[April,P2] + inventory[May,P2]
   = 0
 R30: - x[May,P3] + sales[May,P3] - inventory[April,P3] + inventory[May,P3]
   = 0
 R31: - x[May,P4] + sales[May,P4] - inventory[April,P4] + inventory[May,P4]
   = 0
 R32: - x[May,P5] + sales[May,P5] - inventory[April,P5] + inventory[May,P5]
   = 0
 R33: - x[May,P6] + sales[May,P6] - inventory[April,P6] + inventory[May,P6]
   = 0
 R34: - x[May,P7] + sales[May,P7] - inventory[April,P7] + inventory[May,P7]
   = 0
 R35: - x[June,P1] + sales[June,P1] - inventory[May,P1]
   + inventory[June,P1] = 0
 R36: - x[June,P2] + sales[June,P2] - inventory[May,P2]
   + inventory[June,P2] = 0
 R37: - x[June,P3] + sales[June,P3] - inventory[May,P3]
   + inventory[June,P3] = 0
 R38: - x[June,P4] + sales[June,P4] - inventory[May,P4]
   + inventory[June,P4] = 0
 R39: - x[June,P5] + sales[June,P5] - inventory[May,P5]
   + inventory[June,P5] = 0
 R40: - x[June,P6] + sales[June,P6] - inventory[May,P6]
   + inventory[June,P6] = 0
 R41: - x[June,P7] + sales[June,P7] - inventory[May,P7]
   + inventory[June,P7] = 0
 R42: inventory[June,P1] = 50
 R43: inventory[June,P2] = 50
 R44: inventory[June,P3] = 50
 R45: inventory[June,P4] = 50
 R46: inventory[June,P5] = 50
 R47: inventory[June,P6] = 50
 R48: inventory[June,P7] = 50
 R49: inventory[January,P1] <= 100
 R50: inventory[January,P2] <= 100
 R51: inventory[January,P3] <= 100
 R52: inventory[January,P4] <= 100
 R53: inventory[January,P5] <= 100
 R54: inventory[January,P6] <= 100
 R55: inventory[January,P7] <= 100
 R56: inventory[February,P1] <= 100
 R57: inventory[February,P2] <= 100
 R58: inventory[February,P3] <= 100
 R59: inventory[February,P4] <= 100
 R60: inventory[February,P5] <= 100
 R61: inventory[February,P6] <= 100
 R62: inventory[February,P7] <= 100
 R63: inventory[March,P1] <= 100
 R64: inventory[March,P2] <= 100
 R65: inventory[March,P3] <= 100
 R66: inventory[March,P4] <= 100
 R67: inventory[March,P5] <= 100
 R68: inventory[March,P6] <= 100
 R69: inventory[March,P7] <= 100
 R70: inventory[April,P1] <= 100
 R71: inventory[April,P2] <= 100
 R72: inventory[April,P3] <= 100
 R73: inventory[April,P4] <= 100
 R74: inventory[April,P5] <= 100
 R75: inventory[April,P6] <= 100
 R76: inventory[April,P7] <= 100
 R77: inventory[May,P1] <= 100
 R78: inventory[May,P2] <= 100
 R79: inventory[May,P3] <= 100
 R80: inventory[May,P4] <= 100
 R81: inventory[May,P5] <= 100
 R82: inventory[May,P6] <= 100
 R83: inventory[May,P7] <= 100
 R84: inventory[June,P1] <= 100
 R85: inventory[June,P2] <= 100
 R86: inventory[June,P3] <= 100
 R87: inventory[June,P4] <= 100
 R88: inventory[June,P5] <= 100
 R89: inventory[June,P6] <= 100
 R90: inventory[June,P7] <= 100
 R91: inventory[June,P1] = 50
 R92: inventory[June,P2] = 50
 R93: inventory[June,P3] = 50
 R94: inventory[June,P4] = 50
 R95: inventory[June,P5] = 50
 R96: inventory[June,P6] = 50
 R97: inventory[June,P7] = 50
 R98: sales[January,P1] <= 500
 R99: sales[January,P2] <= 1000
 R100: sales[January,P3] <= 300
 R101: sales[January,P4] <= 300
 R102: sales[January,P5] <= 800
 R103: sales[January,P6] <= 200
 R104: sales[January,P7] <= 100
 R105: sales[February,P1] <= 600
 R106: sales[February,P2] <= 500
 R107: sales[February,P3] <= 200
 R108: sales[February,P4] <= 0
 R109: sales[February,P5] <= 400
 R110: sales[February,P6] <= 300
 R111: sales[February,P7] <= 150
 R112: sales[March,P1] <= 300
 R113: sales[March,P2] <= 600
 R114: sales[March,P3] <= 0
 R115: sales[March,P4] <= 0
 R116: sales[March,P5] <= 500
 R117: sales[March,P6] <= 400
 R118: sales[March,P7] <= 100
 R119: sales[April,P1] <= 200
 R120: sales[April,P2] <= 300
 R121: sales[April,P3] <= 400
 R122: sales[April,P4] <= 500
 R123: sales[April,P5] <= 200
 R124: sales[April,P6] <= 0
 R125: sales[April,P7] <= 100
 R126: sales[May,P1] <= 0
 R127: sales[May,P2] <= 100
 R128: sales[May,P3] <= 500
 R129: sales[May,P4] <= 100
 R130: sales[May,P5] <= 1000
 R131: sales[May,P6] <= 300
 R132: sales[May,P7] <= 0
 R133: sales[June,P1] <= 500
 R134: sales[June,P2] <= 500
 R135: sales[June,P3] <= 100
 R136: sales[June,P4] <= 300
 R137: sales[June,P5] <= 1100
 R138: sales[June,P6] <= 500
 R139: sales[June,P7] <= 60
 R140: 0.5 x[January,P1] + 0.7 x[January,P2] + 0.3 x[January,P5]
   + 0.2 x[January,P6] + 0.5 x[January,P7] <= 1152
 R141: 0.1 x[January,P1] + 0.2 x[January,P2] + 0.3 x[January,P4]
   + 0.6 x[January,P6] <= 768
 R142: 0.2 x[January,P1] + 0.8 x[January,P3] + 0.6 x[January,P7] <= 1152
 R143: 0.05 x[January,P1] + 0.03 x[January,P2] + 0.07 x[January,P4]
   + 0.1 x[January,P5] + 0.08 x[January,P7] <= 384
 R144: 0.01 x[January,P2] + 0.01 x[January,P3] + 0.05 x[January,P4]
   + 0.05 x[January,P7] <= 384
 R145: 0.5 x[February,P1] + 0.7 x[February,P2] + 0.3 x[February,P5]
   + 0.2 x[February,P6] + 0.5 x[February,P7] <= 1536
 R146: 0.1 x[February,P1] + 0.2 x[February,P2] + 0.3 x[February,P4]
   + 0.6 x[February,P6] <= 768
 R147: 0.2 x[February,P1] + 0.8 x[February,P3] + 0.6 x[February,P7] <= 384
 R148: 0.05 x[February,P1] + 0.03 x[February,P2] + 0.07 x[February,P4]
   + 0.1 x[February,P5] + 0.08 x[February,P7] <= 384
 R149: 0.01 x[February,P2] + 0.01 x[February,P3] + 0.05 x[February,P4]
   + 0.05 x[February,P7] <= 384
 R150: 0.5 x[March,P1] + 0.7 x[March,P2] + 0.3 x[March,P5]
   + 0.2 x[March,P6] + 0.5 x[March,P7] <= 1536
 R151: 0.1 x[March,P1] + 0.2 x[March,P2] + 0.3 x[March,P4]
   + 0.6 x[March,P6] <= 768
 R152: 0.2 x[March,P1] + 0.8 x[March,P3] + 0.6 x[March,P7] <= 1152
 R153: 0.05 x[March,P1] + 0.03 x[March,P2] + 0.07 x[March,P4]
   + 0.1 x[March,P5] + 0.08 x[March,P7] <= 0
 R154: 0.01 x[March,P2] + 0.01 x[March,P3] + 0.05 x[March,P4]
   + 0.05 x[March,P7] <= 384
 R155: 0.5 x[April,P1] + 0.7 x[April,P2] + 0.3 x[April,P5]
   + 0.2 x[April,P6] + 0.5 x[April,P7] <= 1536
 R156: 0.1 x[April,P1] + 0.2 x[April,P2] + 0.3 x[April,P4]
   + 0.6 x[April,P6] <= 384
 R157: 0.2 x[April,P1] + 0.8 x[April,P3] + 0.6 x[April,P7] <= 1152
 R158: 0.05 x[April,P1] + 0.03 x[April,P2] + 0.07 x[April,P4]
   + 0.1 x[April,P5] + 0.08 x[April,P7] <= 384
 R159: 0.01 x[April,P2] + 0.01 x[April,P3] + 0.05 x[April,P4]
   + 0.05 x[April,P7] <= 384
 R160: 0.5 x[May,P1] + 0.7 x[May,P2] + 0.3 x[May,P5] + 0.2 x[May,P6]
   + 0.5 x[May,P7] <= 1152
 R161: 0.1 x[May,P1] + 0.2 x[May,P2] + 0.3 x[May,P4] + 0.6 x[May,P6]
   <= 384
 R162: 0.2 x[May,P1] + 0.8 x[May,P3] + 0.6 x[May,P7] <= 1152
 R163: 0.05 x[May,P1] + 0.03 x[May,P2] + 0.07 x[May,P4] + 0.1 x[May,P5]
   + 0.08 x[May,P7] <= 384
 R164: 0.01 x[May,P2] + 0.01 x[May,P3] + 0.05 x[May,P4] + 0.05 x[May,P7]
   <= 384
 R165: 0.5 x[June,P1] + 0.7 x[June,P2] + 0.3 x[June,P5] + 0.2 x[June,P6]
   + 0.5 x[June,P7] <= 1536
 R166: 0.1 x[June,P1] + 0.2 x[June,P2] + 0.3 x[June,P4] + 0.6 x[June,P6]
   <= 768
 R167: 0.2 x[June,P1] + 0.8 x[June,P3] + 0.6 x[June,P7] <= 768
 R168: 0.05 x[June,P1] + 0.03 x[June,P2] + 0.07 x[June,P4] + 0.1 x[June,P5]
   + 0.08 x[June,P7] <= 384
 R169: 0.01 x[June,P2] + 0.01 x[June,P3] + 0.05 x[June,P4]
   + 0.05 x[June,P7] <= 0
Bounds
Generals
 x[January,P1] x[January,P2] x[January,P3] x[January,P4] x[January,P5]
 x[January,P6] x[January,P7] x[February,P1] x[February,P2] x[February,P3]
 x[February,P4] x[February,P5] x[February,P6] x[February,P7] x[March,P1]
 x[March,P2] x[March,P3] x[March,P4] x[March,P5] x[March,P6] x[March,P7]
 x[April,P1] x[April,P2] x[April,P3] x[April,P4] x[April,P5] x[April,P6]
 x[April,P7] x[May,P1] x[May,P2] x[May,P3] x[May,P4] x[May,P5] x[May,P6]
 x[May,P7] x[June,P1] x[June,P2] x[June,P3] x[June,P4] x[June,P5]
 x[June,P6] x[June,P7] sales[January,P1] sales[January,P2]
 sales[January,P3] sales[January,P4] sales[January,P5] sales[January,P6]
 sales[January,P7] sales[February,P1] sales[February,P2] sales[February,P3]
 sales[February,P4] sales[February,P5] sales[February,P6]
 sales[February,P7] sales[March,P1] sales[March,P2] sales[March,P3]
 sales[March,P4] sales[March,P5] sales[March,P6] sales[March,P7]
 sales[April,P1] sales[April,P2] sales[April,P3] sales[April,P4]
 sales[April,P5] sales[April,P6] sales[April,P7] sales[May,P1]
 sales[May,P2] sales[May,P3] sales[May,P4] sales[May,P5] sales[May,P6]
 sales[May,P7] sales[June,P1] sales[June,P2] sales[June,P3] sales[June,P4]
 sales[June,P5] sales[June,P6] sales[June,P7] inventory[January,P1]
 inventory[January,P2] inventory[January,P3] inventory[January,P4]
 inventory[January,P5] inventory[January,P6] inventory[January,P7]
 inventory[February,P1] inventory[February,P2] inventory[February,P3]
 inventory[February,P4] inventory[February,P5] inventory[February,P6]
 inventory[February,P7] inventory[March,P1] inventory[March,P2]
 inventory[March,P3] inventory[March,P4] inventory[March,P5]
 inventory[March,P6] inventory[March,P7] inventory[April,P1]
 inventory[April,P2] inventory[April,P3] inventory[April,P4]
 inventory[April,P5] inventory[April,P6] inventory[April,P7]
 inventory[May,P1] inventory[May,P2] inventory[May,P3] inventory[May,P4]
 inventory[May,P5] inventory[May,P6] inventory[May,P7] inventory[June,P1]
 inventory[June,P2] inventory[June,P3] inventory[June,P4]
 inventory[June,P5] inventory[June,P6] inventory[June,P7]
End
