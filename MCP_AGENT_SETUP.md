# MCP Agent Setup Guide - NSE Dual-Track Stock Analyzer (Daily Automation)

## 🤖 Overview

This guide enables automated daily execution of the **NSE Dual-Track Stock Analyzer** via MCP (Model Context Protocol) agent scheduling. The analyzer tracks Indian equities listed on the National Stock Exchange (NSE).

The repository now also includes an interactive MCP server for stock Q&A against the generated Excel output:
- `stock_excel_mcp_server.py`
- `.vscode/mcp.json`

That server is intended for GitHub Copilot Chat in VS Code when you want to ask natural-language stock questions such as:
- "What does the latest report say about MCX?"
- "Should I buy CYIENT based only on the workbook evidence?"
- "Which query does EQUITASBNK match best?"

The main script now already supports:
- archived run snapshots in `analysis_history/`
- a text run summary in `latest_run_report.txt`
- optional SMTP email delivery driven by `config.json`
- symbol source selection: predefined `config.json` stocks or live AI selection from `.github/skill.md`
- grounded stock Q&A through `python Stock_Agent.py --ask-stock ... --question ...`

That means you can schedule `Stock_Agent.py` directly if you only need daily execution. Use a wrapper only if you want extra orchestration or monitoring.

For interactive IDE usage, prefer the new MCP server instead of the wrapper example below.

---

## Interactive VS Code MCP Setup

### 1. Install MCP Runtime Dependency

```bash
pip install mcp
```

### 2. Generate the Excel Report First

```bash
python Stock_Agent_fixed_v2.py
```

### 3. Use the Included Workspace MCP Config

The repo includes `.vscode/mcp.json` pointing to:

```text
python ${workspaceFolder}/stock_excel_mcp_server.py
```

### 4. Available MCP Tools

- `list_available_stocks`: list stock names present in the latest workbook
- `get_stock_snapshot`: fetch exact workbook-backed fields for a stock
- `answer_stock_question`: answer a natural-language stock question using only the Excel report

### 5. No-Hallucination Design

The Q&A path is intentionally constrained:
- It reads only the `Analysis` sheet from the generated Excel file.
- It keeps recommendation and confidence aligned with stored workbook values.
- If GenAI is enabled, the prompt is evidence-bound and JSON-validated.
- If GenAI is disabled or returns invalid output, the workflow falls back to a deterministic workbook summary.

When using live AI selection in scheduled mode:
- set `symbol_selection.mode` to `live_ai` in `config.json`
- keep `symbol_selection.live_rules_file` pointing to `.github/skill.md`
- keep the screening query in `.github/skill.md` updated to your investment constraints

---

## 📋 Implementation Steps

### Step 1: Create MCP Agent Wrapper

Create a new file: `stock_mcp_agent.py`

```python
"""
MCP Agent wrapper for daily stock analysis execution.
Handles scheduling, error reporting, and result logging.
"""

import subprocess
import json
import logging
from datetime import datetime
from pathlib import Path

# Setup logging
logging.basicConfig(
    filename="mcp_agent.log",
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class StockAnalysisAgent:
    """MCP-compatible stock analysis agent for daily execution."""
    
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.script = "Stock_Agent_fixed_v2.py"
        self.timestamp = datetime.now().isoformat()
    
    def run(self):
        """Execute daily stock analysis and return status."""
        logger.info(f"Starting daily analysis run at {self.timestamp}")
        
        try:
            # Run the main analysis script
            result = subprocess.run(
                ["python", self.script],
                capture_output=True,
                text=True,
                timeout=3600  # 1 hour timeout
            )
            
            if result.returncode == 0:
                logger.info("Analysis completed successfully")
                return self._format_success()
            else:
                logger.error(f"Analysis failed with code {result.returncode}")
                logger.error(f"STDERR: {result.stderr}")
                return self._format_error(result.stderr)
        
        except subprocess.TimeoutExpired:
            logger.error("Analysis timed out after 1 hour")
            return self._format_error("Execution timeout (>1 hour)")
        
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}")
            return self._format_error(str(e))
    
    def _format_success(self):
        """Format successful execution result."""
        config = self._load_config()
        return {
            "status": "success",
            "timestamp": self.timestamp,
            "stocks_processed": len(config.get("stocks", [])),
            "output_file": config.get("export_file", "AI_STOCK_ANALYSIS.xlsx"),
            "log_file": config.get("log_file", "stock_analysis.log"),
            "message": "Daily analysis completed successfully"
        }
    
    def _format_error(self, error_msg):
        """Format error result."""
        return {
            "status": "error",
            "timestamp": self.timestamp,
            "error": error_msg,
            "message": "Daily analysis failed - check logs for details"
        }
    
    def _load_config(self):
        """Load configuration from config.json."""
        try:
            with open(self.config_file, 'r') as f:
                return json.load(f)
        except:
            return {}
    
    def get_latest_results(self):
        """Load and return latest analysis results."""
        try:
            import pandas as pd
            config = self._load_config()
            export_file = config.get("export_file", "AI_STOCK_ANALYSIS.xlsx")
            df = pd.read_excel(export_file)
            
            return {
                "status": "loaded",
                "timestamp": datetime.fromtimestamp(Path(export_file).stat().st_mtime).isoformat(),
                "total_stocks": len(df),
                "top_3_ai_picks": df.nlargest(3, "AI Score")[
                    ["Company Name", "AI Score", "AI Recommendation"]
                ].to_dict(orient="records")
            }
        except Exception as e:
            logger.error(f"Failed to load results: {str(e)}")
            return {"status": "error", "error": str(e)}

def main():
    """Entry point for MCP agent execution."""
    agent = StockAnalysisAgent()
    
    # Run analysis
    result = agent.run()
    
    # Log result
    logger.info(json.dumps(result, indent=2))
    
    # Get latest results summary
    latest = agent.get_latest_results()
    logger.info(f"Latest results: {json.dumps(latest, indent=2)}")
    
    return result

if __name__ == "__main__":
    main()
```

---

### Step 2: Windows Task Scheduler Setup

Create a batch file: `run_daily_analysis.bat`

```batch
@echo off
REM Daily Stock Analysis Runner
REM Scheduled task to run Stock Agent daily at 4:00 PM

cd /d "C:\Users\z003yujx\Documents\AI-Stock-Agent"

REM Optional: Clean up old output files (keep only last 5)
REM forfiles /S /M "AI_STOCK_ANALYSIS_*.xlsx" /D +5 /C "cmd /c del @file"

REM Run Python agent
python stock_mcp_agent.py >> mcp_agent.log 2>&1

REM Exit with status
exit /b %ERRORLEVEL%
```

#### Create Scheduled Task:

1. **Open Task Scheduler** → `taskschd.msc`
2. **Create Basic Task:**
   - Name: `AI Stock Analysis Daily`
   - Description: `Automated daily stock analysis via MCP agent`
   - Trigger: `Daily at 4:00 PM` (after market close at 3:30 PM IST)
   - Action: `Start program` → `C:\path\to\run_daily_analysis.bat`
   - Conditions:
     - Check "Start task only if computer is on AC power"
     - Check "Start the task only if the computer is idle for X minutes"
   - Settings:
     - Check "If the task fails, restart every 15 minutes"
     - Check "Stop the task if it runs longer than 2 hours"

---

### Step 3: Linux/Mac Cron Setup

Add to crontab: `crontab -e`

```bash
# Daily stock analysis at 4:00 PM IST (10:30 AM UTC)
# Run MCP agent
0 16 * * * cd /path/to/AI-Stock-Agent && python stock_mcp_agent.py >> mcp_agent.log 2>&1

# Alternative: Using systemd timer (Linux)
# See systemd service file below
```

#### Systemd Service (Linux):

Create `/etc/systemd/system/stock-analysis.service`:

```ini
[Unit]
Description=AI Stock Analysis Agent
After=network.target

[Service]
Type=oneshot
User=your_username
WorkingDirectory=/path/to/AI-Stock-Agent
ExecStart=/usr/bin/python3 stock_mcp_agent.py
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
```

Create `/etc/systemd/system/stock-analysis.timer`:

```ini
[Unit]
Description=Daily Stock Analysis Trigger
Requires=stock-analysis.service

[Timer]
OnCalendar=*-*-* 16:00:00
Unit=stock-analysis.service
Persistent=true

[Install]
WantedBy=timers.target
```

Enable:
```bash
sudo systemctl daemon-reload
sudo systemctl enable stock-analysis.timer
sudo systemctl start stock-analysis.timer
```

---

### Step 4: Output Organization

Create directories for historical results:

```bash
mkdir -p analysis_results/daily
mkdir -p analysis_results/weekly_archive
mkdir -p analysis_results/monthly_summary
```

Modify `stock_mcp_agent.py` to archive results:

```python
def archive_results(self):
    """Archive analysis results with date prefix."""
    import shutil
    from datetime import datetime
    
    config = self._load_config()
    export_file = config.get("export_file", "AI_STOCK_ANALYSIS.xlsx")
    
    if Path(export_file).exists():
        date_str = datetime.now().strftime("%Y%m%d")
        archive_dir = Path("analysis_results/daily")
        archive_dir.mkdir(parents=True, exist_ok=True)
        
        archive_file = archive_dir / f"analysis_{date_str}.xlsx"
        shutil.copy(export_file, archive_file)
        logger.info(f"Archived results to {archive_file}")
        
        return str(archive_file)
    
    return None
```

---

## 🔔 Notification Setup

### Email Alert on Results

Add to `stock_mcp_agent.py`:

```python
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

def send_email_report(self, results):
    """Send email with top AI picks."""
    import pandas as pd
    
    config = self._load_config()
    if not config.get("email_alerts"):
        return
    
    try:
        # Load results
        df = pd.read_excel(config["export_file"])
        top_picks = df.nlargest(5, "AI Score")[
            ["Company Name", "AI Score", "AI Recommendation", "AI Confidence"]
        ]
        
        # Format email body
        body = f"""
Daily Stock Analysis Report
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

TOP 5 AI PICKS (by AI Score):

{top_picks.to_string(index=False)}

Total stocks analyzed: {len(df)}

Note: Compare AI recommendations with Rule-Based ratings for balanced perspective.

---
This is an automated report from AI Stock Analysis Agent.
        """
        
        # Send email (configure with your email settings)
        msg = MIMEMultipart()
        msg['From'] = config.get("email_from", "your_email@gmail.com")
        msg['To'] = config.get("email_to", "recipient@example.com")
        msg['Subject'] = f"Stock Analysis Report - {datetime.now().strftime('%Y-%m-%d')}"
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach Excel file
        with open(config["export_file"], 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            part.add_header('Content-Disposition', f'attachment; filename= {config["export_file"]}')
            msg.attach(part)
        
        # Send via SMTP
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(config.get("email_from"), config.get("email_password"))
            server.send_message(msg)
        
        logger.info(f"Email report sent to {config.get('email_to')}")
    
    except Exception as e:
        logger.error(f"Failed to send email: {str(e)}")
```

Update `config.json`:

```json
{
  "email_alerts": true,
  "email_from": "your_email@gmail.com",
  "email_password": "your_app_password",
  "email_to": "recipient@example.com"
}
```

---

## 📊 Dashboard Integration (Optional)

### REST API for Results

Create `api_server.py`:

```python
from flask import Flask, jsonify
import pandas as pd

app = Flask(__name__)

@app.route('/api/latest-analysis')
def get_latest_analysis():
    """Return latest analysis results as JSON."""
    try:
        df = pd.read_excel('AI_STOCK_ANALYSIS.xlsx')
        return jsonify({
            "status": "success",
            "timestamp": df["Last Updated"].iloc[0],
            "total_stocks": len(df),
            "top_10": df.nlargest(10, "AI Score")[
                ["Company Name", "AI Score", "AI Recommendation", "Rule-Based Rating"]
            ].to_dict(orient="records")
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route('/api/stock/<symbol>')
def get_stock_analysis(symbol):
    """Return analysis for specific stock."""
    try:
        df = pd.read_excel('AI_STOCK_ANALYSIS.xlsx')
        stock = df[df['Company Name'] == symbol.upper()]
        
        if stock.empty:
            return jsonify({"status": "not_found"}), 404
        
        return jsonify({
            "status": "success",
            "data": stock.iloc[0].to_dict()
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
```

---

## 🛠️ Maintenance

### Weekly Review

```bash
# View latest 10 execution logs
tail -100 mcp_agent.log | grep -E "(success|error|completed)"

# Check file sizes
ls -lah analysis_results/daily/*.xlsx | tail -5

# Archive weekly (every Sunday)
# tar -czf analysis_results/weekly_archive/week_$(date +%Y%m%d).tar.gz analysis_results/daily/
```

### Monthly Cleanup

```bash
# Delete analysis results older than 30 days
find analysis_results/daily -name "*.xlsx" -mtime +30 -delete

# Keep archive for 90 days
find analysis_results/weekly_archive -name "*.tar.gz" -mtime +90 -delete
```

---

## ✅ Monitoring & Alerts

### Health Check Script

Create `health_check.py`:

```python
import os
from datetime import datetime, timedelta

def check_analysis_health():
    """Verify daily analysis is running on schedule."""
    
    # Check if today's analysis exists
    today = datetime.now().strftime("%Y%m%d")
    today_file = f"analysis_results/daily/analysis_{today}.xlsx"
    
    if not os.path.exists(today_file):
        print(f"ALERT: No analysis found for {today}")
        return False
    
    # Check file is recent (within last 24 hours)
    file_time = datetime.fromtimestamp(os.path.getmtime(today_file))
    if datetime.now() - file_time > timedelta(hours=24):
        print(f"ALERT: Analysis file is stale (from {file_time})")
        return False
    
    print(f"OK: Analysis current as of {file_time}")
    return True

if __name__ == "__main__":
    check_analysis_health()
```

Run as hourly cron job for monitoring:
```bash
0 * * * * python health_check.py >> health_check.log 2>&1
```

---

## 📈 Performance Tuning

### Parallel Processing (Optional)

For faster execution with multiple stocks, use `concurrent.futures`:

```python
from concurrent.futures import ThreadPoolExecutor, as_completed

def analyze_parallel(stocks, config, max_workers=5):
    """Analyze multiple stocks in parallel."""
    results = []
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(analyze, sym, config): sym for sym in stocks}
        
        for future in as_completed(futures):
            try:
                result = future.result()
                results.append(result)
            except Exception as e:
                logger.error(f"Error analyzing {futures[future]}: {e}")
    
    return results
```

---

## 🔐 Security Best Practices

1. **Protect Credentials:**
   - Use environment variables for email passwords
   - Never commit credentials to version control
   - Rotate API tokens regularly

2. **Log Rotation:**
   ```bash
   # Rotate logs weekly
   0 0 * * 0 logrotate -f /etc/logrotate.d/stock-analysis
   ```

3. **File Permissions:**
   ```bash
   chmod 700 stock_mcp_agent.py
   chmod 600 config.json
   ```

---

## 🚀 Deployment Checklist

- [ ] `config.json` configured with correct stock symbols
- [ ] `stock_mcp_agent.py` deployed and tested
- [ ] Task Scheduler / Cron job configured
- [ ] Email alerts configured (optional)
- [ ] Log directory created and writable
- [ ] Archive directory created
- [ ] Health check script deployed (optional)
- [ ] Monitoring alerts configured (optional)
- [ ] First test run completed successfully
- [ ] Results verified in output file

---

## 📞 Troubleshooting

| Issue | Solution |
|-------|----------|
| Task doesn't run | Check Task Scheduler logs; verify Python path |
| Email not sending | Verify SMTP credentials; enable "less secure apps" for Gmail |
| Analysis takes too long | Reduce stock count; use parallel processing |
| Memory issues | Split stocks into batches; run on machine with more RAM |
| API rate limiting | Increase delay in `Stock_Agent.py`; use API keys |

---

**Version:** 1.0 (MCP Agent Setup)  
**Last Updated:** 2025-01-15
