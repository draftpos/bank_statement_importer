from odoo import models, fields, api

class AccountJournal(models.Model):
    _inherit = 'account.journal'

    def action_universal_import_wizard(self):
        return {
            'name': 'Import Statement (Universal)',
            'type': 'ir.actions.act_window',
            'res_model': 'universal.import.bank.statement',
            'view_mode': 'form',
            'target': 'new',
            'context': {'default_journal_id': self.id}
        }
