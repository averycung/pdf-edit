import sys  # For GUI
import praw # Reddit API
import pandas as pd # For manipulating data
from collections import Counter
from datetime import datetime 
from dotenv import load_dotenv
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QTextEdit, QPushButton, QMessageBox
) # For GUI
import matplotlib.pyplot as plt # For bar graph

# Load environment variables
load_dotenv()
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
user_agent = os.getenv("USER_AGENT")

# Reddit API setup
reddit = praw.Reddit(
    client_id=client_id,
    client_secret=client_secret,
    user_agent=user_agent
)

class RedditTrackerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reddit Keyword Tracker")
        self.setGeometry(100, 100, 600, 400)
        layout = QVBoxLayout()

        # CSS
        self.setStyleSheet("""              
            QWidget {
                background-color: #f5f5f4;    
                color: #000080;
            }
            
            QLabel {
                font-family: Segoe UI;
                color: #000080; 
                font-weight: bold;
                font-size: 15px;       
            }
            
            QTextEdit {
                border: 2px solid #000080;
                border-radius:5px
            }
                           
            QPushButton {
                background-color: #000080;
                color: #f5f5f4;
                font-family: Segoe UI;
                font-weight: bold;
                font-size: 20px;
                border-radius: 5px;
                padding: 5px 10px;               
            }
                           
            QPushButton:hover {
                background-color: #000040;
                color: #f5f5f4;
            }
                           """)

        # Subreddit input
        layout.addWidget(QLabel("Enter subreddits (comma-separated):"))
        self.subreddits_input = QTextEdit()
        self.subreddits_input.setFixedHeight(40)
        layout.addWidget(self.subreddits_input)

        # Keyword input
        layout.addWidget(QLabel("Enter keywords (comma-separated):"))
        self.keywords_input = QTextEdit()
        self.keywords_input.setFixedHeight(40)
        layout.addWidget(self.keywords_input)

        layout.addStretch()

        # Track button
        self.track_button = QPushButton("Track Keywords")
        self.track_button.setFixedHeight(50)
        self.track_button.clicked.connect(self.track_keywords)
        layout.addWidget(self.track_button)

        # Set layout
        self.setLayout(layout)

    
    def track_keywords(self):
        # Track subreddits and keywords
        subreddits = [s.strip() for s in self.subreddits_input.toPlainText().split(",")]
        keywords = [k.strip().lower() for k in self.keywords_input.toPlainText().split(",")]

        posts = []
        for sub in subreddits:
            try:
                for post in reddit.subreddit(sub).new(limit=300):  # Past 300 posts
                    posts.append(post.title.lower() + " " + post.selftext.lower())
            # Error handling in case subreddit can't be found
            except Exception as e:
                print(f"Skipping subreddit '{sub}': {e}")

        counter = Counter()
        for kw in keywords:
            counter[kw] = sum(kw in post for post in posts)

        df = pd.DataFrame(counter.items(), columns=["Keyword", "Mentions"])
        # Track trending keywords if they are mentioned >40 times in a day
        df["Trending"] = df["Mentions"] > 40
        df["Date"] = datetime.now().strftime("%Y-%m-%d")
        df = df.sort_values(by="Mentions", ascending=False)

        df.to_csv("reddit_mentions.csv", index=False)
        df.to_csv("trend_log.csv", mode="a", header=not os.path.exists("trend_log.csv"), index=False)

        df.head(100).plot(kind="bar", x="Keyword", y="Mentions", title="Top Keyword Mentions", legend=False) 
        
        # Styling of chart
        top_df = df.head(100)
        bars = plt.bar(top_df["Keyword"], top_df["Mentions"], color="#000080", edgecolor="#000000")
        plt.title("Top Keyword Mentions on Reddit", fontfamily='Garamond', fontsize=12, weight='bold', color="#000000")
        plt.xlabel("Keyword", fontsize=11, fontfamily='Garamond')
        plt.ylabel("Mentions", fontsize=11, fontfamily='Garamond')
        plt.xticks(rotation=45, ha='right', fontfamily='Garamond', fontsize=10)
        plt.yticks(fontsize=10, fontfamily='Garamond')
        plt.tight_layout()
        plt.savefig("keyword_visualization.png")
        plt.show()

        trending = df[df["Trending"] == True]
        if not trending.empty:
            message = "Trending keywords:\n" + trending[["Keyword", "Mentions"]].to_string(index=False)
        else:
            message = "No trending keywords found."

        QMessageBox.information(self, "Tracking Complete", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RedditTrackerApp()
    window.show()
    sys.exit(app.exec_())
