import xlwt as xlwt

class DownoadReport(APIView):

    def get(self, request):
        
        # using a dummy user model with three fields name,age,date_of_birth
        all_users = users.objects.all()
        
        # values_list() returns tuples those are immutable so maping 
        # the objects into list of lists because i want to change values of data
        all_users = list(map(list, (all_users.values_list('name', 'age', 'date_of_birth'))))
        
        # list to define column names
        columns = ['Name', 'Age', 'Date Of Birth']
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="user.xls"'
        
        # creating excel workbook and add a sheet to it
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Users data')

        row_num = 0
        # custom styles for headers
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num], font_style)

        # Sheet body, remaining rows
        sheet_style = xlwt.XFStyle()
        # Customise your date patteren other wise it will provide an ISO standard date
        font_style.num_format_str = 'dd-mm-yy'
        
        for row in all_users:
        # Here i decide if the user is under 18 it will replace the value 
        # with "Not Adult" else it will replace it with "Adult"
            if row[1] < 18:
                row[1] = 'Not Adult'
            else:
                row[1] = 'Adult'
            row_num += 1
            for col_num in range(len(row)):
                ws.write(row_num, col_num, row[col_num], sheet_style)

        wb.save(response)
        return response
