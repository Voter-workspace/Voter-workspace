from storage.supabase import supabase

def add_work_entry(data):
    return (
        supabase
        .table("work_entries")
        .insert(data)
        .execute()
    )

def get_all_entries():
    response = (
        supabase
        .table("work_entries")
        .select("*")
        .order("id", desc=True)
        .execute()
    )
    return response.data

def update_work_entry(entry_id, data):
    return (
        supabase
        .table("work_entries")
        .update(data)
        .eq("id", entry_id)
        .execute()
    )

def delete_work_entry(entry_id):
    return (
        supabase
        .table("work_entries")
        .delete()
        .eq("id", entry_id)
        .execute()
    )