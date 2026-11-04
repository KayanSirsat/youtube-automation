# yt_intel: competitor and niche intelligence for the asset-economics channel

Automates the research behind `Luxury_Asset_Economics_YouTube_Market_Entry.xlsx`:
competitor metrics, outlier videos, evergreen tracking, comment mining, hook
analysis and the saturation scan, all from the YouTube Data API v3.

## 1. Setup (about 10 minutes)

1. Python 3.10 or newer.
2. `pip install -r requirements.txt`
3. Get an API key:
   - Google Cloud Console, create a project.
   - APIs and Services, Library, enable **YouTube Data API v3**.
   - Credentials, Create credentials, API key.
   - Restrict the key to YouTube Data API v3.
4. Set the key as an environment variable (do not commit it):
   - macOS/Linux: `export YT_API_KEY="your-key"`
   - Windows: `setx YT_API_KEY "your-key"` then open a new terminal.
5. Check the cost before the first run: `python yt_intel.py --estimate`

## 2. Commands

| Command | What it does | Approx. quota units |
|---|---|---|
| `python yt_intel.py` | Full run: channels, uploads, comments, saturation scan | about 2,600 |
| `python yt_intel.py --snapshot-only` | Channels and uploads only, for daily tracking | about 72 |
| `python yt_intel.py --skip-search` | Everything except the scan | about 150 |
| `python yt_intel.py --estimate` | Prints the estimate, makes no calls | 0 |

Costs per call: channels, playlistItems, videos and commentThreads are 1 unit;
search is 100 units. New projects usually get 10,000 units per day; check
yours in the console under Quotas. `quota_budget_per_run` in config.json stops
a run before it overspends, and whatever was collected is still exported.

## 3. Schedule

Daily snapshots are what turn this into a Social Blade replacement: from the
second run on, every channel and video gets **Recent Views per Day**, and
videos get **Recent vs Lifetime Daily Rate** (a value near or above 1 on a
video older than a year means it is still evergreen).

- Daily, 06:00: `--snapshot-only`
- Weekly, Monday 07:00: full run

macOS/Linux crontab (`crontab -e`):

```
0 6 * * * cd /path/to/yt_intel && /usr/bin/python3 yt_intel.py --snapshot-only >> run.log 2>&1
0 7 * * 1 cd /path/to/yt_intel && /usr/bin/python3 yt_intel.py >> run.log 2>&1
```

Windows Task Scheduler: create a basic task, action "Start a program",
program `python`, arguments `yt_intel.py --snapshot-only`, start in the
yt_intel folder. Repeat for the weekly full run.

## 4. Outputs

Written to `exports/`: a dated workbook, `latest.xlsx` (always the newest run,
point Power BI here), and one CSV per sheet in `exports/csv/`. History lives in
`yt_intel.db` (SQLite); keep it, it is what produces the trend columns.

| Sheet | Answers |
|---|---|
| Run_Log | What ran, quota used, anything skipped or not found |
| Channels | Subs, views, videos, avg views per video, subs per 1K views, median length, uploads per week, Shorts share, recent views per day |
| Videos | Every sampled upload with format, title pattern, outlier multiple, evergreen rate |
| Top_Videos | Best videos per channel by multiple of channel median, with an Outlier flag |
| Comments | Top comments on each channel's best videos, tagged by theme |
| Comment_Themes | Share of comments per theme per channel (questions, AI complaints, corrections, topic requests, praise, voice, visuals) |
| Requested_Topics | Topics viewers ask for, ranked by comment likes: your idea backlog |
| Comment_Bigrams | Most frequent word pairs per channel |
| Transcript_Hooks | First 15 seconds text, words per minute, numbers in first 60 seconds, first sponsor mention (only if enabled) |
| Saturation_Scan | Every video found per asset class, now with views, channel and age |
| Saturation_Summary | Per asset class: videos, distinct channels, median views, share under 150 days, videos per channel |
| Saturation_Publishers | Channels in the scan; "Likely Template Publisher" = covers 3 or more asset classes |

## 5. Power BI

1. Get Data, Excel workbook, `exports/latest.xlsx` (or Folder, `exports/csv`).
2. Load Channels, Videos, Top_Videos, Comments, Comment_Themes, Saturation_Scan, Saturation_Summary, Saturation_Publishers.
3. Relationships: `Videos[Channel ID]` many to one `Channels[Channel ID]`; `Comments[Video ID]` many to one `Videos[Video ID]`.
4. Starter measures:

```
Avg Views per Video = DIVIDE(SUM(Channels[Total Views]), SUM(Channels[Videos]))
Outlier Rate = DIVIDE(CALCULATE(COUNTROWS(Top_Videos), Top_Videos[Outlier] = "Yes"), COUNTROWS(Top_Videos))
Evergreen Videos = CALCULATE(COUNTROWS(Videos), Videos[Age Days] >= 365, Videos[Recent vs Lifetime Daily Rate] >= 0.5)
AI Complaint Share = CALCULATE(AVERAGE(Comment_Themes[Share of Comments]), Comment_Themes[Theme] = "AI or slop complaint")
Scan Share Under 150 Days = DIVIDE(CALCULATE(COUNTROWS(Saturation_Scan), Saturation_Scan[Title Match] = "Yes", Saturation_Scan[Under 150 Days] = "Yes"), CALCULATE(COUNTROWS(Saturation_Scan), Saturation_Scan[Title Match] = "Yes"))
```

5. Refresh after each run. Column names are stable between runs.

## 6. Editing config.json

- `channels`: handles (`@name`) or channel IDs (`UC...`). Handles were taken from Social Blade, vidIQ and OutlierKit URLs; anything not found is listed in Run_Log.
- `max_videos_per_channel`: newest uploads sampled per channel (200 = 8 units per channel).
- `comments`: how many top videos per channel and comments per video.
- `transcripts.enabled`: see section 7 before turning on.
- `saturation_scan.assets`: add asset classes; `terms` must appear in the title and `intent_regex` must match for a video to count.

## 7. Limits and compliance

- **Sample bias:** uploads are read newest first, so an old all-time hit beyond the newest 200 is missed (matters for Alux). Raise `max_videos_per_channel` for small channels.
- **Shorts** are detected by length (180 seconds or less), which is a heuristic.
- **Themes** are keyword rules, good for direction, not precision. `exports/csv/Comments.csv` has one comment per row with a stable ID, so it can go straight into your existing hybrid VADER plus LLM sentiment pipeline for better labels.
- **Search results** are YouTube's relevance ranking for one query, not a census of the niche.
- **API policies:** YouTube's developer policies include rules on how long API data may be stored and when it must be refreshed. Read the current YouTube API Services Terms and Developer Policies before keeping long history in `yt_intel.db`, and prune old rows if required. Commenter names are deliberately not stored.
- **Transcripts** use `youtube-transcript-api`, an unofficial library that reads captions outside the official API. It can break, can get your IP rate-limited, and may conflict with YouTube's terms, so it is off by default. Turn it on knowingly, keep `top_videos_per_channel` small and the delay at 2 seconds or more.
- **Not tested against the live API:** I built and ran this end to end against mocked API responses (pagination, disabled comments, missing channels, snapshot deltas, export), because my sandbox cannot reach googleapis.com. Run `--snapshot-only` first and check Run_Log.
