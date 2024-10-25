from django.utils.deprecation import MiddlewareMixin


class CookieMiddleware(MiddlewareMixin):
    def process_response(self, request, response):
        if request.path == '/final/':

            if request.session.get('file1'):
                response.set_cookie('file1', request.session.get('file1'))
            if request.session.get('file2'):
                response.set_cookie('file2', request.session.get('file2'))
            if request.session.get('file3'):
                response.set_cookie('file3', request.session.get('file3'))
            if request.session.get('date'):
                response.set_cookie('date', request.session.get('date'))

        return response
