# Measure some strings:
words = ['Ram', 'Rambo', 'Ramesh']

for i in words:
    print(i, len(i))

# Create a sample collection

users = {'Ram': 'Active', 'Rambo': 'inactive', 'Ramesh': 'Suspended', 'Rahul': "Active"}

for user, status in users.copy().items():
    if status == "inactive":
        del users[user]
print(users)

# Strategy:  Create a new collection
active_users = {}
for user, status in users.items():
    if status == 'Active':
        active_users[user] = status
print(active_users)

# 4.3. The range() Function
# If you do need to iterate over a sequence of numbers, the built-in function range() comes in handy.
# It generates arithmetic progressions:

for i in range(5):
    print(i)
print("/n")
for j in range(5, 10):
    print(j)
print("/n")
for k in range(0, 10, 3):
    print(k)
print("/n")

for i in range(-10, -100, -30):
    print(i)
print("/n")
