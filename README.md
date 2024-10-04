# Crewdle Mist SharePoint External Storage Connector

## Introduction

The Crewdle Mist SharePoint External Storage Connector for Node is a seamless solution built to enable smooth integration with SharePoint for external file storage. This connector allows Node applications to effortlessly store, retrieve, and manage files directly within SharePoint, providing a secure and scalable storage option. With easy setup and strong support for SharePoint’s file management features, it is an ideal solution for developers looking to extend their applications’ storage capabilities into the SharePoint ecosystem while maintaining data integrity and accessibility.

## Getting Started

Before diving in, ensure you have installed the [Crewdle Mist SDK](https://www.npmjs.com/package/@crewdle/web-sdk).

## Installation

```bash
npm install @crewdle/mist-connector-sharepoint
```

## Usage

```TypeScript
import { SharepointExternalStorageConnector } from '@crewdle/mist-connector-sharepoint';

// Create a new SDK instance
const sdk = await SDK.getInstance('[VENDOR ID]', '[ACCESS TOKEN]', {
  externalStorageConnectors: new Map([
    ['sharepoint', SharepointExternalStorageConnector]
  ])
});
```

## Need Help?

Reach out to support@crewdle.com or raise an issue in our repository for any assistance.

## Join Our Community

For an engaging discussion about your specific use cases or to connect with fellow developers, we invite you to join our Discord community. Follow this link to become a part of our vibrant group: [Join us on Discord](https://discord.gg/XJ3scBYX).
