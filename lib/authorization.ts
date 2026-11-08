import { auth } from '@/lib/auth';
import { getDb } from '@/lib/mongodb';
import { isAllowedEmailDomain } from '@/lib/email-domain';

export async function getAuthorizedEmail() {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();
  if (!email || !isAllowedEmailDomain(email)) return null;

  const db = await getDb();
  const member = await db.collection('members').findOne({ emailNormalized: email, active: true });
  return member ? email : null;
}
