/**
 * Скрипт для реєстрації команд в Discord
 * Використовується для розгортання slash-команд
 */
type Mode = 'global' | 'guild' | 'both';
interface DeployOptions {
    dry?: boolean;
    mode?: Mode;
    guildId?: string;
}
declare function deployCommands(options?: DeployOptions): Promise<void>;
export { deployCommands };
//# sourceMappingURL=deployCommands.d.ts.map