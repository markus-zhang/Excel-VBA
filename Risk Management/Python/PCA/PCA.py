import numpy as np
import math
from numpy import linalg as LA
from scipy.stats import norm

### Implementation of a class structure for the above code ###
class cInterestTSVaR:
	'A base class to calculate VaR based on Term Structure with PCA for reduction'
	dataraw = np.empty([2, 2])
	datanum = np.empty([2, 2])
	dataevector = np.empty([2, 2])
	dataevalue = np.empty([1,1])
	datapcevector = np.empty([2, 2])
	datapcevalue = np.empty([1,1])
	evaluethresh = 0

	def __init__(self, filename, delim, withheader, threshold):
		self.dataraw = np.genfromtxt(filename, delimiter = delim, skip_header = withheader)
		self.evaluethresh = threshold

	def GetNum(self):
		if math.isnan(self.dataraw[1, 1]):
			self.datanum = self.dataraw
		else:
			self.datanum = [ row[1:] for row in self.dataraw]

	def GetEigen(self):
		self.GetNum()
		self.datartn = np.diff(self.datanum, axis = 0) * 100	#returns in bps
		self.datacov = np.cov(self.datartn, rowvar = 0)
		self.dataevalue, self.dataevector = LA.eig(self.datacov)
		#Need to multiply all the eigenvectors, except the first one with -1
		self.dataevector[:, 1:] *= -1

	def GetPC(self):
		evaluetemp = self.dataevalue / sum(self.dataevalue)
		total = 0
		for i in range(0, evaluetemp.size):
			if total >= self.evaluethresh:
				self.datapcevalue = self.dataevalue[0:i]
				self.datapcevector = self.dataevector[:, :i]
				return
			total += evaluetemp[i]

	def GetVaR(self, pv01, siglevel, rhorizon):
		#Get Net sensitivities by multiplying PV01 vector (e.g. 1*60) with the PC eigenmatrix (e.g. 60*3 for 3 eigenvectors)
		#Get P&L Var by adding up the multiplication of every net sensitivity of an eigenvector with its eigenvalue
		#P&L Vol is square root of P&L var multiplied by square root of risk horizon, this is your sigma in the formula
		#VaR is calculated by -NORMSINV(1-significance level) * P&L Vol * SQRT(h / 250) in Excel
		netsen = np.matmul(pv01 * 1000, self.datapcevector) 
		print(netsen)
		plvol = np.sqrt(np.dot(np.square(netsen), self.datapcevalue))
		print(plvol)
		var = -1 * norm.ppf(siglevel) * plvol * np.sqrt(rhorizon)
		print(var)


	def Dump(self, precision=3):
		print('The Eigenvector is:\n')
		print(np.around(self.dataevector, precision))
		print('The Eigenvalues are:\n')
		print(np.around(self.dataevalue, precision))

### Test ###

temp = cInterestTSVaR('LIBOR.csv', ',', 1, 0.99)
temp.GetEigen()
temp.GetPC()
pv01 = np.genfromtxt('PV01.csv', delimiter = ',', skip_header = 0)
temp.GetVaR(pv01, 0.01, 10)