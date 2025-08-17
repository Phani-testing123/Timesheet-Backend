import os
import boto3

s3 = boto3.client('s3', endpoint_url=os.getenv('S3_ENDPOINT') or None, region_name=os.getenv('S3_REGION'))

def upload_bytes_and_get_url(path: str, content: bytes, content_type: str) -> str:
    bucket = os.getenv('S3_BUCKET')
    if not bucket:
        # local fallback: write to /tmp and return a placeholder path
        local_path = '/tmp/' + path.replace('/', '_')
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, 'wb') as f:
            f.write(content)
        return 'file://' + local_path
    s3.put_object(Bucket=bucket, Key=path, Body=content, ContentType=content_type)
    url = s3.generate_presigned_url(
        ClientMethod='get_object',
        Params={'Bucket': bucket, 'Key': path},
        ExpiresIn=60*60*24*7
    )
    return url
