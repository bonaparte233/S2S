"""Context processors for ppt_generator app."""


def user_role(request):
    """Add user role information to all templates."""
    is_developer = False
    
    if request.user.is_authenticated:
        is_developer = (
            request.user.groups.filter(name="开发者").exists() 
            or request.user.is_superuser
        )
    
    return {
        "is_developer": is_developer,
    }

