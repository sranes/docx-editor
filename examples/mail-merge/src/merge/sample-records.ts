import type { MergeRecord } from './merge-engine.ts';

/** Two sample records. They take opposite branches of every rule. */
export const RECORDS: readonly MergeRecord[] = [
  {
    values: {
      FirstName: 'Ada',
      LastName: 'Lovelace',
      CustomerId: 'C-1001',
      City: 'London',
      Balance: '$120.50',
      AccountNo: 'ACC-778812',
      Sender: 'The Billing Team',
      Company: 'Analytical Engines Ltd',
      Closing: 'With best regards,',
    },
    flags: { isVip: true, hasBalance: true, hasItems: true },
    lists: {
      items: [
        { name: 'Punch cards', qty: '200', price: '$40.00' },
        { name: 'Brass gears', qty: '12', price: '$80.50' },
      ],
      payments: [
        { date: '2026-09-01', amount: '$60.00' },
        { date: '2026-09-15', amount: '$40.50' },
      ],
    },
  },
  {
    values: {
      FirstName: 'Alan',
      LastName: 'Turing',
      CustomerId: 'C-1002',
      City: 'Manchester',
      Balance: '$0.00',
      AccountNo: 'ACC-114455',
      Sender: 'The Billing Team',
      Company: 'Analytical Engines Ltd',
      Closing: 'Kind regards,',
    },
    flags: { isVip: false, hasBalance: false, hasItems: false },
    lists: { items: [], payments: [] },
  },
];
