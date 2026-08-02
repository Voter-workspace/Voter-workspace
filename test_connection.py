from supabase import create_client

SUPABASE_URL = "https://nahwsdwzlocezbcukrmv.supabase.co"
SUPABASE_KEY = "sb_publishable_esoz3SCBUza9ufdF5Ia5eg_WD1QJ_2-"

try:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

    response = (
        supabase
        .table("users")
        .select("*")
        .execute()
    )

    print("✅ Connected Successfully")
    print(response.data)

except Exception as e:
    print("❌ Error")
    print(type(e).__name__)
    print(e)