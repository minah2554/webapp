// api/admin-auth.js

export default async function handler(req, res) {
    // Only allow POST requests
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
    }

    const { password } = req.body;

    // Read ADMIN_PASSWORD from environment variables, fallback to 'minah'
    const adminPassword = process.env.ADMIN_PASSWORD || 'minah';

    if (password === adminPassword) {
        return res.status(200).json({ success: true });
    } else {
        return res.status(200).json({ success: false });
    }
}
