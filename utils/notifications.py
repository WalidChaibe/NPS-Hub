# utils/notifications.py
from datetime import datetime, timedelta
from supabase import Client


def get_unread_count(supabase: Client, user_id: str) -> int:
    try:
        res = supabase.table("notifications")\
            .select("id", count="exact")\
            .eq("user_id", user_id)\
            .eq("is_read", False)\
            .execute()
        return res.count or 0
    except Exception:
        return 0


def get_notifications(supabase: Client, user_id: str, limit: int = 20):
    try:
        res = supabase.table("notifications")\
            .select("*")\
            .eq("user_id", user_id)\
            .order("created_at", desc=True)\
            .limit(limit)\
            .execute()
        return res.data or []
    except Exception:
        return []


def mark_all_read(supabase: Client, user_id: str):
    try:
        supabase.table("notifications")\
            .update({"is_read": True})\
            .eq("user_id", user_id)\
            .eq("is_read", False)\
            .execute()
    except Exception:
        pass


def mark_one_read(supabase: Client, notif_id: str):
    try:
        supabase.table("notifications")\
            .update({"is_read": True})\
            .eq("id", notif_id)\
            .execute()
    except Exception:
        pass


def create_notification(supabase: Client, user_id: str, title: str,
                        message: str, notif_type: str = "info",
                        action_plan_id: str = None):
    try:
        payload = {
            "user_id": user_id,
            "title": title,
            "message": message,
            "type": notif_type,
            "is_read": False,
        }
        if action_plan_id:
            payload["action_plan_id"] = action_plan_id
        supabase.table("notifications").insert(payload).execute()
    except Exception:
        pass


def check_action_plan_notifications(supabase: Client):
    """
    Call this once per session on app load.
    Checks all open action plans and creates notifications if:
    - Due date is today or tomorrow (approaching)
    - Due date has passed (overdue)
    Avoids duplicate notifications by checking if one already exists today.
    """
    try:
        # Load all open/in-progress action plans across all pillars
        tables = [
            ("qm_action_plans", "QM"),
            ("am_action_plans", "AM"),
            ("pm_action_plans", "PM"),
            ("hse_action_plans", "HSE"),
            ("fi_action_plans",  "FI"),
            ("et_action_plans",  "ET"),
        ]
        today = datetime.now().date()

        for table, pillar in tables:
            try:
                res = supabase.table(table)\
                    .select("*")\
                    .in_("status", ["Open", "In Progress"])\
                    .execute()
                plans = res.data or []
            except Exception:
                continue

            for plan in plans:
                due_str = plan.get("due_date")
                owner   = plan.get("owner", "")
                action  = plan.get("action", "")[:60]
                plan_id = plan.get("id")
                if not due_str or not owner:
                    continue

                try:
                    due_date = datetime.strptime(due_str[:10], "%Y-%m-%d").date()
                except Exception:
                    continue

                days_left = (due_date - today).days

                # Find the user_id for this owner name
                try:
                    user_res = supabase.table("profiles")\
                        .select("id")\
                        .eq("full_name", owner)\
                        .execute()
                    if not user_res.data:
                        continue
                    user_id = user_res.data[0]["id"]
                except Exception:
                    continue

                # Check if we already sent this notification today
                try:
                    today_start = datetime.combine(today, datetime.min.time()).isoformat()
                    existing = supabase.table("notifications")\
                        .select("id")\
                        .eq("user_id", user_id)\
                        .eq("action_plan_id", plan_id)\
                        .gte("created_at", today_start)\
                        .execute()
                    if existing.data:
                        continue  # already notified today
                except Exception:
                    pass

                if days_left < 0:
                    create_notification(
                        supabase, user_id,
                        title=f"🔴 Overdue — {pillar} Action Plan",
                        message=f'"{action}" was due on {due_date.strftime("%d %b %Y")} and is still open.',
                        notif_type="overdue",
                        action_plan_id=plan_id
                    )
                elif days_left <= 2:
                    create_notification(
                        supabase, user_id,
                        title=f"🟡 Due Soon — {pillar} Action Plan",
                        message=f'"{action}" is due on {due_date.strftime("%d %b %Y")} ({days_left} day{"s" if days_left != 1 else ""} left).',
                        notif_type="warning",
                        action_plan_id=plan_id
                    )

    except Exception:
        pass
