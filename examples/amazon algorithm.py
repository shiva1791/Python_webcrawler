__author__ = 'Shiva'
# Given a number K, find the smallest Fibonacci number that shares a common factor( other than 1 ) with it.
# A number is said to be a common factor of two numbers if it exactly divides both of them.

# code to find factors of a number
class algo:
    def factors(n):
        for i in range(1,  n + 1):
            if n % i == 0:
                print(i)

if __name__ == '__main__':
    algo.factors(12)
