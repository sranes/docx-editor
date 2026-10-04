import { createRoot } from 'react-dom/client';
import { MailMergeApp } from './MailMergeApp';

const container = document.getElementById('app');
if (container) createRoot(container).render(<MailMergeApp />);
