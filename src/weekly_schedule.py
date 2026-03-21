#!/usr/bin/env python3
"""
週間予定メール送信スクリプト

毎週月曜日08:00(JST)に実行し、Google Calendarから1週間分の予定を取得し
フォーマットしたメールをGmail APIで送信する。
"""

import os
import json
import base64
import logging
import sys
from datetime import datetime, timedelta
from email.mime.text import MIMEText

import pytz
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ログ設定
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 定数
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RECIPIENT_EMAIL = "takami.sp@gmail.com"
SENDER_NAME = "高見徹"
JST = pytz.timezone('Asia/Tokyo')
WEEKDAY_JA = ['月', '火', '水', '木', '金', '土', '日']
MAX_RETRY = 3
GOOGLE_SCOPES = [
    'https://www.googleapis.com/auth/calendar.readonly',
    'https://www.googleapis.com/auth/gmail.send'
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 日付ユーティリティ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def get_week_range():
    """
    実行日から7日間の日付範囲を返す（JST）
    Returns: (start_date, end_date) の tuple
    """
    now = datetime.now(JST)
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end = (start + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=0)
    return start, end


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Google API セットアップ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def setup_google_services():
    """
    GitHub SecretsのGOOGLE_TOKEN_JSONを使いGoogle Calendar/Gmail APIを初期化する
    Returns: (calendar_service, gmail_service) の tuple
    """
    logger.info("Google APIサービスを初期化しています...")

    token_json = os.environ.get('GOOGLE_TOKEN_JSON')
    if not token_json:
        raise ValueError("環境変数 GOOGLE_TOKEN_JSON が設定されていません")

    token_data = json.loads(token_json)
    creds = Credentials.from_authorized_user_info(token_data, scopes=GOOGLE_SCOPES)

    calendar_service = build('calendar', 'v3', credentials=creds)
    gmail_service = build('gmail', 'v1', credentials=creds)

    logger.info("Google APIサービスの初期化が完了しました")
    return calendar_service, gmail_service


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# カレンダーイベント取得
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def fetch_calendar_events(calendar_service, start_date, end_date):
    """
    Google Calendarから指定期間のイベントを取得する（リトライあり）
    Returns: イベントのリスト
    """
    logger.info(f"カレンダーイベントを取得中: {start_date.strftime('%Y/%m/%d')} 〜 {end_date.strftime('%Y/%m/%d')}")

    last_error = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            result = calendar_service.events().list(
                calendarId='primary',
                timeMin=start_date.isoformat(),
                timeMax=end_date.isoformat(),
                singleEvents=True,
                orderBy='startTime',
                timeZone='Asia/Tokyo'
            ).execute()

            events = result.get('items', [])
            logger.info(f"取得したイベント数: {len(events)}件")
            return events

        except HttpError as e:
            last_error = e
            logger.warning(f"カレンダー取得 試行{attempt}/{MAX_RETRY} 失敗: {e}")
            if attempt < MAX_RETRY:
                import time
                time.sleep(2 ** attempt)  # 指数バックオフ

    raise last_error


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# メール本文フォーマット
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def format_event_detail(event):
    """
    1イベントの詳細テキストを生成する
    Returns: フォーマット済みテキスト（複数行）
    """
    lines = []
    start = event.get('start', {})
    end = event.get('end', {})

    # 時刻 or 終日
    if 'dateTime' in start:
        start_dt = datetime.fromisoformat(start['dateTime'])
        end_dt = datetime.fromisoformat(end['dateTime'])
        time_str = f"　{start_dt.strftime('%H:%M')}〜{end_dt.strftime('%H:%M')}"
        lines.append(f"{time_str}　{event.get('summary', '（タイトルなし）')}")
    else:
        lines.append(f"　（終日）　{event.get('summary', '（タイトルなし）')}")

    # 説明（備考）
    description = event.get('description', '').strip()
    if description:
        for line in description.split('\n'):
            if line.strip():
                lines.append(f"　{line.strip()}")

    return '\n'.join(lines)


def format_email_body(events, start_date, end_date):
    """
    1週間分のメール本文を生成する
    Returns: フォーマット済みメール本文
    """
    # 日付ごとにイベントを分類
    events_by_date = {}
    current = start_date
    while current.date() <= end_date.date():
        events_by_date[current.date()] = []
        current += timedelta(days=1)

    for event in events:
        start = event.get('start', {})
        if 'dateTime' in start:
            event_date = datetime.fromisoformat(start['dateTime']).date()
        else:
            event_date = datetime.fromisoformat(start.get('date', '')).date()

        if event_date in events_by_date:
            events_by_date[event_date].append(event)

    # ヘッダー
    year = start_date.year
    start_str = f"{start_date.month}月{start_date.day}日"
    end_str = f"{end_date.month}月{end_date.day}日"

    body_lines = [
        f"{SENDER_NAME}の１週間の予定（{year}年{start_str}〜{end_str}）",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    ]

    # 各日の予定
    for date, day_events in sorted(events_by_date.items()):
        weekday = WEEKDAY_JA[date.weekday()]
        body_lines.append(f"\n■ {date.month}月{date.day}日（{weekday}）")

        if day_events:
            for event in day_events:
                body_lines.append(format_event_detail(event))
        else:
            body_lines.append("　予定なし")

    # フッター
    body_lines.append("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    body_lines.append("以上です。")

    return '\n'.join(body_lines)


def build_subject(start_date, end_date):
    """
    メール件名を生成する
    例: 高見徹の１週間の予定（2026/3/21〜2026/3/27）
    """
    start_str = f"{start_date.year}/{start_date.month}/{start_date.day}"
    end_str = f"{end_date.year}/{end_date.month}/{end_date.day}"
    return f"{SENDER_NAME}の１週間の予定（{start_str}〜{end_str}）"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 送信先検証・メール送信
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def validate_recipient(recipient):
    """
    送信先アドレスが正しいか検証する（誤送信防止）
    送信先が takami.sp@gmail.com 以外の場合は例外を発生させる
    """
    if recipient != RECIPIENT_EMAIL:
        raise ValueError(
            f"[送信先検証エラー] 期待値={RECIPIENT_EMAIL}, 実際={recipient}\n"
            f"送信先が正しくないためメール送信を中止しました。"
        )
    logger.info(f"✓ 送信先アドレス検証OK: {recipient}")


def send_email(gmail_service, subject, body, recipient):
    """
    Gmail APIでメールを送信する（リトライあり）
    送信前に必ず送信先アドレスを検証する
    """
    # ━ 送信先の検証（必須）━
    validate_recipient(recipient)

    logger.info(f"メールを送信します...")
    logger.info(f"  宛先: {recipient}")
    logger.info(f"  件名: {subject}")

    # メールの作成
    message = MIMEText(body, 'plain', 'utf-8')
    message['to'] = recipient
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

    last_error = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            result = gmail_service.users().messages().send(
                userId='me',
                body={'raw': raw}
            ).execute()
            logger.info(f"✓ メール送信成功: message_id={result.get('id')}")
            return result

        except HttpError as e:
            last_error = e
            logger.warning(f"メール送信 試行{attempt}/{MAX_RETRY} 失敗: {e}")
            if attempt < MAX_RETRY:
                import time
                time.sleep(2 ** attempt)

    raise last_error


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# メイン処理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def main():
    logger.info("=" * 50)
    logger.info("週間予定メール送信 開始")
    logger.info("=" * 50)

    exit_code = 0

    try:
        # ステップ1: Google APIサービスの初期化
        logger.info("[STEP 1/5] Google APIサービスを初期化")
        calendar_service, gmail_service = setup_google_services()

        # ステップ2: 日付範囲の取得
        logger.info("[STEP 2/5] 対象期間を確認")
        start_date, end_date = get_week_range()
        logger.info(f"  対象期間: {start_date.strftime('%Y/%m/%d')} 〜 {end_date.strftime('%Y/%m/%d')}")

        # ステップ3: カレンダーイベントの取得
        logger.info("[STEP 3/5] カレンダーイベントを取得")
        events = fetch_calendar_events(calendar_service, start_date, end_date)

        # ステップ4: メール件名・本文の生成
        logger.info("[STEP 4/5] メール内容を生成")
        subject = build_subject(start_date, end_date)
        body = format_email_body(events, start_date, end_date)
        logger.info(f"  件名: {subject}")
        logger.info(f"  本文プレビュー（先頭200文字）:\n{body[:200]}...")

        # ステップ5: メール送信（送信先検証込み）
        logger.info("[STEP 5/5] メールを送信")
        send_email(gmail_service, subject, body, RECIPIENT_EMAIL)

        logger.info("=" * 50)
        logger.info("週間予定メール送信 完了")
        logger.info("=" * 50)

    except ValueError as e:
        logger.error(f"[検証エラー] {e}")
        exit_code = 1

    except HttpError as e:
        logger.error(f"[Google APIエラー] {e}")
        exit_code = 2

    except Exception as e:
        logger.error(f"[予期しないエラー] {e}", exc_info=True)
        exit_code = 3

    return exit_code


if __name__ == '__main__':
    sys.exit(main())
