__author__ = 'Shiva'

class fibonacci:

    def fibo_using_for_loop(self, n):
        a,b = 0,1
        for i in range(n):
            print (a, end = ",")
            a,b = b, a+b

if __name__ == '__main__':

    n = int(input("How many terms? "))

    fi = fibonacci()
    fi.fibo_using_for_loop(n)
