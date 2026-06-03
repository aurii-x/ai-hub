'use client';

import { Box, SimpleGrid, Text, Card, CardBody, Heading } from '@chakra-ui/react';

export default function DashboardPage() {
  return (
    <Box p="24px">
      <Heading mb="24px">Santhosh OS Dashboard</Heading>

      <SimpleGrid columns={{ base: 1, md: 2, xl: 4 }} spacing="20px">
        <Card>
          <CardBody>
            <Text fontSize="sm" color="gray.500">Tasks</Text>
            <Text fontSize="2xl" fontWeight="bold">Todoist</Text>
            <Text mt="8px">Today’s tasks will show here.</Text>
          </CardBody>
        </Card>

        <Card>
          <CardBody>
            <Text fontSize="sm" color="gray.500">Email</Text>
            <Text fontSize="2xl" fontWeight="bold">Gmail</Text>
            <Text mt="8px">Unread and important emails will show here.</Text>
          </CardBody>
        </Card>

        <Card>
          <CardBody>
            <Text fontSize="sm" color="gray.500">Calendar</Text>
            <Text fontSize="2xl" fontWeight="bold">Google Calendar</Text>
            <Text mt="8px">Upcoming events will show here.</Text>
          </CardBody>
        </Card>

        <Card>
          <CardBody>
            <Text fontSize="sm" color="gray.500">Automations</Text>
            <Text fontSize="2xl" fontWeight="bold">Scripts</Text>
            <Text mt="8px">Apps Script, AppleScript, and BTT controls.</Text>
          </CardBody>
        </Card>
      </SimpleGrid>
    </Box>
  );
}