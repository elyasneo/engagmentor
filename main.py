from ast import literal_eval
from  igramscraper.instagram import Instagram # pylint: disable=import-error
from  igramscraper.model.media import Media # pylint: disable=import-error
import getpass
import stdiomask
import os
import xlsxwriter
from datetime import datetime
clear = lambda: os.system('cls')
def deleteContent(pfile):
    pfile.seek(0)
    pfile.truncate()

# If account is public you can query Instagram without auth
instagram = Instagram()
while True:
    file = open('cred.txt',"r+")
    up=file.read().splitlines()
    if len(up)<2 or (not up[0] or not up[1]):
        username = input('username: ')
        password = stdiomask.getpass('pass: ')
    else:
        username=up[0]
        password=up[1]
    print("try to login...")
    try:
        instagram.with_credentials(username,password)
        instagram.login()
        deleteContent(file)
        file.write(f"{username}\n{password}")
        file.close()
        break
    except Exception as e:
        deleteContent(file)
        file.close()
        print(e)
        raise e


clear()

while True:

    try:
        instagram.login()
    except Exception as e:
        file = open('cred.txt',"r+")
        deleteContent(file)
        file.close()
        raise e

    targetUsername=input("Enter Target Username: ")
    numOfPost=input("Number Of Posts: ")
    # For getting information about account you don't need to auth:
    print('please wait this may take a few minutes...')
    account = instagram.get_account(targetUsername)
    posts = instagram.get_medias(targetUsername,int(numOfPost))

    followerCount = account.followed_by_count 

    def engagementRatio(post):
        return float(post.likes_count/followerCount)

    def printPost(posts):
        book = xlsxwriter.Workbook(f"{targetUsername}.xlsx")

        infoSheet=book.add_worksheet("Info")
        infoSheet.set_column(0,1,20)

        infoSheet.write(0,0,"Username")
        infoSheet.write(0,1,account.username)

        infoSheet.write(1,0,"Name")
        infoSheet.write(1,1,account.full_name)

        infoSheet.write(2,0,"Followers")
        infoSheet.write(2,1,account.followed_by_count)

        infoSheet.write(3,0,"Following")
        infoSheet.write(3,1,account.follows_count)

        infoSheet.write(4,0,"Posts")
        infoSheet.write(4,1,account.media_count)

        infoSheet.write(5,0,"AVG. Likes")
        
        infoSheet.write(6,0,"AVG. Comments")
        

        engagementSheet = book.add_worksheet("Engagement")
        engagementSheet.set_column(3,3,100)
        engagementSheet.set_column(0,2,20)
        engagementSheet.write(0,0,"percent")
        engagementSheet.write(0,1,"likes")
        engagementSheet.write(0,2,"comments")
        engagementSheet.write(0,3,"link")
        sumLikes=0
        sumComments=0
        sumVideoViews=0
        for i in range(0,len(posts)):
            sumLikes += posts[i].likes_count
            sumComments += posts[i].comments_count
            sumVideoViews += posts[i].video_views
            engagementRatioPercent = round(100.0 * posts[i].likes_count / float(followerCount), 2)
            postLink = posts[i].link
            engagementSheet.write(i+1,0,engagementRatioPercent)
            engagementSheet.write(i+1,1,posts[i].likes_count)
            engagementSheet.write(i+1,2,posts[i].comments_count)
            engagementSheet.write(i+1,3,postLink)

        if len(posts)>0:
            infoSheet.write(5,1,round(sumLikes/len(posts), 2))
            infoSheet.write(6,1,round(sumComments/len(posts),2))
        else:
            infoSheet.write(5,1,0)
            infoSheet.write(6,1,0)
        book.close()


    # Available fields

    posts.sort(key=engagementRatio,reverse=True)
    while True:
        try:
            printPost(posts)
            break
        except Exception as e:
            input("\nPlease close Excel!! And press y: ")
    
    print(f"\n{targetUsername}.xlsx Created")
    print()
