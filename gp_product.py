from odoo import api, fields, models, _
from odoo.exceptions import UserError, ValidationError
from odoo.tools.misc import xlsxwriter
from odoo.tools.misc import xlwt
from xlwt import easyxf
import io
import base64

class GpProductReport(models.TransientModel):
    _name = "gp.product.report"
    _description = "GP Report"

    from_date = fields.Date(String="From Date", required=True)
    to_date = fields.Date(String="To Date", required=True)
    report_summary_file = fields.Binary('Payment Summary Report')
    file_name = fields.Char(string='File Name')
    flag = fields.Boolean(string='Flag Name',default=False)

    @api.multi
    def action_report(self,data):
    	print("1111111111111111")
        workbook = xlwt.Workbook()
        column_heading_style = easyxf('font:height 200;font:bold True;align: horiz center;borders: top thin,bottom thin,left thin,right thin,top_color black, bottom_color black, right_color  black, left_color black;')
        column_heading_style_up = easyxf('font:height 200;font:bold True;align: horiz center;borders: top thin,bottom thin,left thin,right thin;')
        column_heading_style2 = easyxf('font:height 200;font:bold False;align: horiz left;')
        column_heading_style3 = easyxf('font:height 200;font:bold True;align: horiz center;font: color red;borders: top thin,bottom thin,left thin,right thin,top_color black, bottom_color black, right_color  black, left_color black;')
        dateFormat = xlwt.XFStyle()
        dateFormat.num_format_str = 'dd/mm/yyyy/'
        style2 = xlwt.easyxf(num_format_str='#,##0.00')
        decimal_style = xlwt.XFStyle()
        decimal_style.num_format_str = '0.0000'
        number_style = xlwt.easyxf("", "#,###.00")
        style_string = "font: bold on; borders: top thin,bottom thin,left thin,right thin"
        style = xlwt.easyxf(style_string)

        worksheet = workbook.add_sheet('Productwise GP Statement')
        # worksheet.write_merge(1, 0, 1, 10, _('Productwise GP Statement From'), column_heading_style) 
        worksheet.write(1, 3, _('Productwise GP Statement From '+str(self.from_date)+' To '+str(self.to_date)), column_heading_style) 
        worksheet.write(3, 0, _('Product'), column_heading_style) 
        worksheet.write(3, 1, _('Product Category'), column_heading_style)     
        worksheet.write(3, 2, _('Quantity'), column_heading_style)
        worksheet.write(3, 3, _('Sales'), column_heading_style)
        worksheet.write(3, 4, _('Cost Of Sales'), column_heading_style)
        worksheet.write(3, 5, _('GP Amt'), column_heading_style)
        worksheet.write(3, 6, _('GP %'), column_heading_style)
        inv,inv_lines,=[],[]
        inv = self.env['account.invoice'].search([('date_invoice','>=',self.from_date),('date_invoice','<=',self.to_date),('state','in',['open','paid']),('type','=','out_invoice')])
        if inv:
            inv_lines = self.env['account.invoice.line'].search([('invoice_id','in',inv.ids)])
        prd =[]
        if inv_lines:
            for i in inv_lines:
                if i.product_id not in prd:
                    prd.append(i.product_id)
        row = 4
        for p in prd:
            p_lines =self.env['account.invoice.line'].search([('invoice_id','in',inv.ids),('product_id','=',p.id)])
            qty_tot,cost,sale,avg_cost,avg_sale,gp_per =0,0,0,0,0,0  
            for pl in p_lines:
                qty_tot = qty_tot + pl.sale_line_ids.product_uom_qty
                cost = cost + pl.sale_line_ids.purchase_price
                sale = sale + pl.sale_line_ids.price_unit
            if qty_tot > 0:
                avg_cost =  cost/qty_tot
                avg_sale =  sale/qty_tot  
            if avg_sale:    
                gp_per = round((avg_sale - avg_cost)/avg_sale*100,2)    
            worksheet.write(row, 0, p.name)   
            worksheet.write(row, 1, p.categ_id.complete_name)    
            worksheet.write(row, 2, qty_tot)   
            worksheet.write(row, 3, round(avg_sale,2))
            worksheet.write(row, 4, round(avg_cost,2)) 
            worksheet.write(row, 5, round((avg_sale - avg_cost),2))    
            worksheet.write(row, 6, gp_per)    
            row = row + 1
        fp = io.BytesIO()
        workbook.save(fp)
        excel_file = base64.encodestring(fp.getvalue())
        self.report_summary_file = excel_file
        self.file_name = 'Productwise GP Report.xls'
        self.flag = True
        fp.close()
        return {
                'view_mode': 'form',
                'res_id': self.id,
                'res_model': 'gp.product.report',
                'view_type': 'form',
                'type': 'ir.actions.act_window',
                'context': self.env.context,
                'target': 'new',
                   }


class GpProductSalespersonReport(models.TransientModel):
    _name = "gp.salesman.report"
    _description = "GP Report"

    from_date = fields.Date(String="From Date", required=True)
    to_date = fields.Date(String="To Date", required=True)
    salesman = fields.Many2many('res.users',String="Salesman")
    report_summary_file = fields.Binary('Payment Summary Report')
    file_name = fields.Char(string='File Name')
    flag = fields.Boolean(string='Flag Name',default=False)

    @api.multi
    def action_report(self,data):
        workbook = xlwt.Workbook()
        column_heading_style = easyxf('font:height 200;font:bold True;align: horiz center;borders: top thin,bottom thin,left thin,right thin,top_color black, bottom_color black, right_color  black, left_color black;')
        column_heading_style_up = easyxf('font:height 200;font:bold True;align: horiz center;borders: top thin,bottom thin,left thin,right thin;')
        column_heading_style2 = easyxf('font:height 200;font:bold False;align: horiz left;')
        column_heading_style3 = easyxf('font:height 200;font:bold True;align: horiz center;font: color red;borders: top thin,bottom thin,left thin,right thin,top_color black, bottom_color black, right_color  black, left_color black;')
        dateFormat = xlwt.XFStyle()
        dateFormat.num_format_str = 'dd/mm/yyyy/'
        style2 = xlwt.easyxf(num_format_str='#,##0.00')
        decimal_style = xlwt.XFStyle()
        decimal_style.num_format_str = '0.0000'
        number_style = xlwt.easyxf("", "#,###.00")
        style_string = "font: bold on; borders: top thin,bottom thin,left thin,right thin"
        style = xlwt.easyxf(style_string)

        worksheet = workbook.add_sheet('SalesPerson Wise GP Statement')
        # worksheet.write_merge(1, 0, 1, 10, _('Productwise GP Statement From'), column_heading_style) 
        worksheet.write(1, 3, _('Productwise GP Statement From '+str(self.from_date)+' To '+str(self.to_date)), column_heading_style) 
        worksheet.write(3, 0, _('Salesman'), column_heading_style) 
        worksheet.write(3, 1, _('Product'), column_heading_style) 
        worksheet.write(3, 2, _('Product Category'), column_heading_style)     
        worksheet.write(3, 3, _('Quantity'), column_heading_style)
        worksheet.write(3, 4, _('Sales'), column_heading_style)
        worksheet.write(3, 5, _('Cost Of Sales'), column_heading_style)
        worksheet.write(3, 6, _('GP Amt'), column_heading_style)
        worksheet.write(3, 7, _('GP %'), column_heading_style)
        row = 4
        inv,inv_lines,=[],[]
        if self.salesman:
            salesman = self.salesman
        else:
            salesman = self.env['res.users'].search([])
        for sale in salesman:
            inv = self.env['account.invoice'].search([('date_invoice','>=',self.from_date),('date_invoice','<=',self.to_date),('state','in',['open','paid']),('user_id','=',sale.id),('type','=','out_invoice')])
            if inv:
                inv_lines = self.env['account.invoice.line'].search([('invoice_id','in',inv.ids)])
            prd =[]
            if inv_lines:
                for i in inv_lines:
                    if i.product_id not in prd:
                        prd.append(i.product_id)
            flag=0
            for p in prd:
                p_lines =self.env['account.invoice.line'].search([('invoice_id','in',inv.ids),('product_id','=',p.id)])
                qty_tot =0 
                for pl in p_lines:
                    qty_tot = qty_tot + pl.sale_line_ids.product_uom_qty
                    if qty_tot > 0:
                        flag= 1
            if flag == 1:
                worksheet.write(row, 0, sale.name)
                row = row + 1
            # for p in prd:
                # p_lines =self.env['account.invoice.line'].search([('invoice_id','in',inv.ids),('product_id','=',p.id)])
                # qty_tot,cost,sale,avg_cost,avg_sale,gp_per =0,0,0,0,0,0  
                # for pl in p_lines:
                    # qty_tot = qty_tot + pl.quantity
                    
            for p in prd:
                p_lines =self.env['account.invoice.line'].search([('invoice_id','in',inv.ids),('product_id','=',p.id)])
                qty_tot,cost,sale,avg_cost,avg_sale,gp_per =0,0,0,0,0,0  
                for pl in p_lines:
                    qty_tot = qty_tot + pl.sale_line_ids.product_uom_qty
                    cost = cost + pl.sale_line_ids.purchase_price
                    sale = sale + pl.sale_line_ids.price_unit
                if qty_tot > 0:
                    avg_cost =  cost/qty_tot
                    avg_sale =  sale/qty_tot  
                if avg_sale:    
                    gp_per = round((avg_sale - avg_cost)/avg_sale*100,2)  
                if qty_tot > 0:
                    worksheet.write(row, 1, p.name)   
                    worksheet.write(row, 2, p.categ_id.complete_name)
                    worksheet.write(row, 3, qty_tot)   
                    worksheet.write(row, 4, round(avg_sale))
                    worksheet.write(row, 5, round(avg_cost)) 
                    worksheet.write(row, 6, round((avg_sale - avg_cost),2))
                    worksheet.write(row, 7, gp_per)    
                    row = row + 1
            
        fp = io.BytesIO()
        workbook.save(fp)
        excel_file = base64.encodestring(fp.getvalue())
        self.report_summary_file = excel_file
        self.file_name = 'Salesmanwise GP Report.xls'
        self.flag = True
        fp.close()
        return {
                'view_mode': 'form',
                'res_id': self.id,
                'res_model': 'gp.salesman.report',
                'view_type': 'form',
                'type': 'ir.actions.act_window',
                'context': self.env.context,
                'target': 'new',
                   }
    
            #
