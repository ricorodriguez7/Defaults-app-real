import NextAuth from 'next-auth';
import GoogleProvider from 'next-auth/providers/google';
import { getDb } from './mongodb';
import { isAllowedEmailDomain } from './email-domain';

export const { handlers, auth } = NextAuth({
  providers: [
    GoogleProvider({
      clientId: process.env.AUTH_GOOGLE_ID || '',
      clientSecret: process.env.AUTH_GOOGLE_SECRET || '',
    }),
  ],
  session: {
    strategy: 'jwt',
  },
  pages: {
    signIn: '/',
  },
  callbacks: {
    async signIn({ user }) {
      if (!user.email) return false;
      const email = user.email.trim().toLowerCase();
      if (!isAllowedEmailDomain(email)) return false;
      const db = await getDb();
      const member = await db.collection('members').findOne({ emailNormalized: email, active: true });
      return Boolean(member);
    },
    async session({ session }) {
      if (session.user) {
        session.user.email = session.user.email?.toLowerCase() || '';
      }
      return session;
    },
  },
});

export const { GET, POST } = handlers;
