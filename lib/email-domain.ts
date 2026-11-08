const allowedEmailDomain = (process.env.ALLOWED_EMAIL_DOMAIN || 'lbsim.ac.in')
  .trim()
  .toLowerCase()
  .replace(/^@/, '');

export function isAllowedEmailDomain(email: string) {
  return email.trim().toLowerCase().endsWith(`@${allowedEmailDomain}`);
}
