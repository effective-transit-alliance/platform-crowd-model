#This is a recursive platform clearance calculator.
#c is total area of VCEs in sqft. Use 1500.
#k0 is passengers aboard train on arrival
#u is rate of train egress in pax/second; use 1 pax/second/door
#You input desired t1, time elapsed after arrival
#w is total width of upstairs VCEs in feet.
#egress is instantaneous rate of platform egress
c = 10000
k0 = 1200
u = 40
t1 = 150
w = 35
egress = 0
def integratel(f,a,b,n):
    a1 = float(a)
    b1 = float(b)
    y = 0
    for i in range(0,n):
        y = y + f(a1 + (b1-a1)*i/n)
    return y*(b1-a1)/n
n = 500

def trainloadfn(k,u,t):
    return max(0,k-u*t)



for i in range (1,t1):
    def platloadfn(k, u, t): #number of passengers on platform, determines space per pax
        if trainloadfn(k,u,t) > 0:
            return max(1,min(k, (u * t - egress * t)))
        else:
            return max(1,k - egress * t)
    def crowdfn(k, u, t, c): #space per pax, determines egress
        return c / platloadfn(k, u, t)
    egress = 2/60*w*(111 *crowdfn(k0, u, i-1, c)-162)/(crowdfn(k0, u, i-1, c)**2)
    instplatload = platloadfn(k0,u,i)
    instcrowding = crowdfn(k0,u,i,c)
    print('Platform load at t = ' + str(i) + ' is ' + str(instplatload) + ' passengers. Each passenger has ' + str(instcrowding)  + ' sf of space. The egress rate is ' + str(egress) + ' pax/sec.')
