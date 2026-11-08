import { auth } from '@/lib/auth';
import { getDb } from '@/lib/mongodb';

export async function getAuthorizedEmail() {
  const session = await auth();
  const email = session?.user?.email?.trim().toLowerCase();
  if (!email) return null;

  const db = await getDb();
  const member = await db.collection('members').findOne({ emailNormalized: email, active: true });
  return member ? email : null;
}
