from django.conf import settings
from django.contrib.auth import logout
from django.shortcuts import redirect
from django.utils import timezone


class IdleTimeoutMiddleware:
    """
    Enforce the inactivity timeout server-side.

    Reads `settings.INACTIVITY_TIMEOUT_SECONDS` (default 600 = 10 min)
    and compares it against `request.session['_last_activity']`.

    Called on every authenticated request. If the user has been idle
    longer than the configured threshold, they are logged out and
    redirected to the login page.
    """

    def __init__(self, get_response):
        self.get_response = get_response
        self.timeout = getattr(settings, 'INACTIVITY_TIMEOUT_SECONDS', 600)

    def __call__(self, request):
        user = getattr(request, 'user', None)

        # Only check authenticated requests
        if user is not None and user.is_authenticated:
            now  = timezone.now()
            last = request.session.get('_last_activity')

            if last:
                try:
                    last_dt = timezone.datetime.fromisoformat(last)
                except (ValueError, TypeError):
                    last_dt = None

                if last_dt:
                    idle_seconds = (now - last_dt).total_seconds()
                    if idle_seconds > self.timeout:
                        logout(request)
                        request.session.flush()
                        # Send them to the login page.
                        # `LOGIN_URL` in settings is a URL name
                        # (`control_dashboard:login`), so use reverse.
                        from django.urls import reverse
                        login_url = settings.LOGIN_URL
                        if not login_url.startswith('/'):
                            try:
                                login_url = reverse(login_url)
                            except Exception:
                                login_url = '/login/'
                        return redirect(login_url)

            # Refresh the timestamp for the next request
            request.session['_last_activity'] = now.isoformat()

        return self.get_response(request)