import { DeviceCodeCredential } from '@azure/identity';

const clientId = process.env.AZURE_CLIENT_ID;
const tenantId = process.env.AZURE_TENANT_ID || 'common';

if (!clientId) {
  console.error('Set AZURE_CLIENT_ID env var first.');
  console.error('For single-tenant work apps, also set AZURE_TENANT_ID to your Directory (tenant) ID.');
  process.exit(1);
}

console.log('Testing with clientId:', clientId);
console.log('Tenant:', tenantId);

const credential = new DeviceCodeCredential({
  clientId,
  tenantId,
  userPromptCallback: (info) => {
    console.log('\n=== DEVICE CODE INFO (full object) ===');
    console.log(JSON.stringify(info, null, 2));
    console.log('=======================================\n');
    console.log('Open:', info.verificationUri);
    console.log('Enter code:', info.userCode);
    console.log('\nWaiting for you to authenticate in browser...\n');
  },
});

try {
  const token = await credential.getToken(['User.Read']);
  console.log('SUCCESS! Token expires:', new Date(token.expiresOnTimestamp));
} catch (err) {
  console.error('FAILED:');
  console.error('  name:', err.name);
  console.error('  message:', err.message);
  console.error('  code:', err.code);
  console.error('  errorCode:', err.errorCode);
  console.error('  subError:', err.subError);
  console.error('  statusCode:', err.statusCode);
  if (err.responseBody) console.error('  responseBody:', err.responseBody);
  console.error('\nFull error object:');
  console.error(err);
}
