# 🚀 ShareEasy - Setup Guide

## What is ShareEasy?

ShareEasy is an AI-powered file sharing and format conversion app — built to eliminate the friction of "I have it in JPG but they need it as a DOCX" situations.

**The problem it solves:**
Normally, if someone sends you a file in the wrong format, you have to:
1. Manually download a converter tool
2. Convert the file yourself
3. Re-send it to the recipient

**With ShareEasy, it's just one prompt:**
> *"Send these files to Priya in DOCX format"*

The app automatically converts the file (JPG → DOCX, PDF → TXT, XLSX → CSV, and many more) and generates a shareable download link for the recipient — no manual steps needed.

**How it works:**
- **Sender** uploads a file and tells the AI bot the target format and recipient
- **AI bot** (powered by Claude) understands the instruction and triggers the conversion
- **Receiver** gets a link to view and download the converted file instantly

**Supported formats:** PDF · DOCX · TXT · PNG · JPG · WEBP · XLSX · CSV

**Future vision:**
> ShareEasy is designed to work as a plugin inside WhatsApp, Microsoft Teams, Slack, and other messaging platforms — so file conversion and sharing becomes a natural part of any conversation, without ever leaving the app.

---

---

## 📋 PART 1: Test Locally (On Your Laptop)

### Step 1: Install Node.js
1. Go to: https://nodejs.org
2. Download the **LTS version** (should say "Recommended for most users")
3. Run the installer
4. Keep clicking "Next" until it finishes
5. **Verify it worked:**
   - Press `Windows key`, type `cmd`, press Enter
   - Type: `node --version`
   - You should see something like `v20.x.x`
   - If you see a version number, you're good! ✅

### Step 2: Set Up the Project
1. Extract the `ShareEasy` folder to your Desktop (or anywhere you want)
2. Open VS Code
3. Click **File → Open Folder**
4. Select the `ShareEasy` folder
5. You should see all the files in the left sidebar

### Step 3: Get Your Claude API Key (for the AI bot feature)
1. Go to: https://console.anthropic.com
2. Sign in (or create a free account)
3. Click **"Get API Keys"**
4. Click **"Create Key"**
5. Copy the key (starts with `sk-ant-...`)

### Step 4: Configure the API Key
1. In VS Code, find the file `.env.example`
2. Right-click it → **Rename** → change name to `.env` (remove the `.example`)
3. Open the `.env` file
4. Replace `your_api_key_here` with your actual API key:
   ```
   ANTHROPIC_API_KEY=sk-ant-your-actual-key-here
   PORT=3000
   ```
5. Save the file (Ctrl + S)

### Step 5: Install Dependencies
1. In VS Code, click **Terminal → New Terminal** (or press Ctrl + `)
2. A terminal window appears at the bottom
3. Type this command and press Enter:
   ```
   npm install
   ```
4. Wait 1-2 minutes. You'll see lots of text scrolling — that's normal!
5. When it says "added X packages" and shows a prompt again, you're done ✅

### Step 6: Run the App
1. In the same terminal, type:
   ```
   npm start
   ```
2. You should see:
   ```
   ✅  ShareEasy is running!
   👉  Open browser → http://localhost:3000
   ```
3. Open your browser (Chrome, Edge, Firefox — any works)
4. Go to: `http://localhost:3000`
5. **You should see the ShareEasy app!** 🎉

### Step 7: Test It
1. Click **"Upload a file"** and pick any document/image
2. Choose a recipient name
3. Pick a target format (e.g., PDF)
4. Click **"Send File"**
5. If it works, you'll see a success message with a share link!

**Stopping the app:**
- In the terminal, press `Ctrl + C`
- The app stops. To restart, just run `npm start` again.

---

## 🌍 PART 2: Deploy Online (Render.com)

Once your app works locally, deploy it so others can use it!

### Step 1: Create a GitHub Account (if you don't have one)
1. Go to: https://github.com
2. Click **"Sign up"**
3. Create a free account

### Step 2: Install Git
1. Go to: https://git-scm.com
2. Download and install
3. Keep all default settings during installation
4. Verify: Open `cmd` and type `git --version` — should show a version

### Step 3: Upload Your Code to GitHub
1. Go to: https://github.com/new
2. **Repository name:** `shareeasy`
3. Set to **Private** (so only you can see the code)
4. Click **"Create repository"**
5. You'll see a page with instructions — **keep this tab open**

6. Back in VS Code terminal, type these commands **one by one**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/shareeasy.git
   git push -u origin main
   ```
   (Replace `YOUR-USERNAME` with your actual GitHub username)

7. Enter your GitHub username and password when asked
8. Refresh the GitHub page — you should see your files uploaded! ✅

### Step 4: Deploy to Render
1. Go to: https://render.com
2. Click **"Get Started"** → Sign up with GitHub
3. After signing in, click **"New +"** → **"Web Service"**
4. Click **"Connect a repository"** → find `shareeasy` → **Connect**
5. Fill in the form:
   - **Name:** `shareeasy` (or anything you want)
   - **Environment:** `Node`
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Instance Type:** `Free`
6. Click **"Advanced"** → **"Add Environment Variable"**
   - **Key:** `ANTHROPIC_API_KEY`
   - **Value:** (paste your Claude API key)
7. Click **"Create Web Service"**

### Step 5: Wait for Deployment
- Render will build your app (takes 2-3 minutes)
- When you see **"Live"** with a green dot, it's ready!
- Click the URL at the top (looks like `https://shareeasy-xyz.onrender.com`)
- **Your app is now online!** 🌍

### Step 6: Share the Link
- Send the URL to anyone
- They can upload files and use the bot without installing anything!

---

## 🔧 Troubleshooting

**App won't start locally:**
- Make sure you ran `npm install` first
- Check that Node.js is installed: `node --version`
- Make sure you're in the right folder in terminal

**"Cannot find module" error:**
- Run `npm install` again

**Bot doesn't respond:**
- Check your `.env` file has the correct API key
- Make sure the key starts with `sk-ant-`

**Files won't convert:**
- Some conversions need extra tools (we use basic ones for now)
- Check the terminal for error messages

**Render deployment failed:**
- Check you added the `ANTHROPIC_API_KEY` environment variable
- Make sure your GitHub repo is public or Render is connected to it

---

## 📁 Project Structure

```
file-share-bot/
├── server.js          ← Backend (Node.js/Express)
├── package.json       ← Dependencies list
├── .env              ← Your API key (keep this SECRET!)
├── public/
│   ├── index.html    ← Main web page
│   ├── style.css     ← Styles
│   └── app.js        ← Frontend logic
├── uploads/          ← Temporary upload storage
└── converted/        ← Converted files
```

---

## 🎓 What You Learned

Even though I wrote the code, you now understand:
- ✅ How web apps work (frontend talks to backend via API)
- ✅ How Node.js servers handle file uploads
- ✅ How AI APIs work (sending requests to Claude)
- ✅ How to deploy apps online for free
- ✅ How Git and GitHub work for version control

**Next steps to learn coding yourself:**
- Try changing colors in `style.css` and see what happens
- Modify text in `index.html`
- Add console.log() in `app.js` to see what's happening
- Read through the comments in the code — they explain everything!

---

## 🆘 Need Help?

If something doesn't work:
1. Check the terminal for error messages
2. Make sure all steps were followed in order
3. Google the exact error message (seriously — this is what developers do!)
4. Ask me for help and paste the error message

---

## 📝 License

This is your project! Do whatever you want with it.
- Customize it
- Deploy it
- Share it
- Learn from it
- Build something even better!

Good luck! 🚀
