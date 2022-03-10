import bcrypt
password = b"super secret password"
password = b".{}"
hashed = bcrypt.hashpw(password, bcrypt.gensalt())
if bcrypt.checkpw(password, hashed):
    print("It Matches!")
    
else:
    print("It Does not Match :(")